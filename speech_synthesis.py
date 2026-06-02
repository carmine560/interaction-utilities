"""Speech process management and control utilities."""

from multiprocessing import Pipe, Process
import time

import win32com.client

from core_utilities.errors import ProcessStateError

READY_TIMEOUT_SECONDS = 5
JOIN_TIMEOUT_SECONDS = 5
TERMINATE_TIMEOUT_SECONDS = 1


class SpeechManager:
    """Manage the speech capabilities and text of a speech system."""

    def __init__(self):
        """Initialize a new SpeechManager instance."""
        self._is_ready = False
        self._can_speak = True
        self._speech_text_queue = []

    def is_ready(self):
        """Get whether the speech system is initialized and ready."""
        return self._is_ready

    def set_ready(self, value):
        """Set whether the speech system is initialized and ready."""
        self._is_ready = value

    def can_speak(self):
        """Get the speech capability of the system."""
        return self._can_speak

    def set_can_speak(self, can_speak):
        """Set the speech capability of the system."""
        self._can_speak = can_speak

    def get_speech_text(self):
        """Get the speech text of the system."""
        return self._speech_text_queue[0] if self._speech_text_queue else ""

    def set_speech_text(self, text):
        """Set the speech text of the system."""
        if text:
            self._speech_text_queue.append(text)
        else:
            self._speech_text_queue.clear()

    def pop_speech_text(self):
        """Pop the next speech text from the system."""
        if self._speech_text_queue:
            return self._speech_text_queue.pop(0)
        return ""


def start_speaking_process(
    speech_manager,
    voice_name=None,
    speech_rate=None,
    ready_timeout=READY_TIMEOUT_SECONDS,
):
    """Start a new process for speaking."""
    startup_error_receiver, startup_error_sender = Pipe(duplex=False)
    speaking_process = Process(
        target=start_speaking,
        args=(speech_manager, voice_name, speech_rate, startup_error_sender),
    )
    speaking_process.start()
    startup_error_sender.close()
    deadline = time.monotonic() + ready_timeout
    while not speech_manager.is_ready():
        # First, check whether the child explicitly reported a startup failure.
        try:
            startup_error_available = startup_error_receiver.poll()
        except (BrokenPipeError, EOFError):
            startup_error_available = False
        if startup_error_available:
            startup_error = startup_error_receiver.recv()
            speaking_process.join(timeout=TERMINATE_TIMEOUT_SECONDS)
            startup_error_receiver.close()
            raise ProcessStateError(
                "Speech process failed to start: " f"{startup_error}"
            )
        # Next, fail fast if the child exited before it could become ready.
        # Re-check the pipe after join in case the error arrived just before
        # exit.
        if getattr(speaking_process, "exitcode", None) is not None:
            speaking_process.join(timeout=TERMINATE_TIMEOUT_SECONDS)
            try:
                startup_error_available = startup_error_receiver.poll()
            except (BrokenPipeError, EOFError):
                startup_error_available = False
            if startup_error_available:
                startup_error = startup_error_receiver.recv()
                startup_error_receiver.close()
                raise ProcessStateError(
                    "Speech process failed to start: " f"{startup_error}"
                )
            startup_error_receiver.close()
            raise ProcessStateError(
                "Speech process exited before becoming ready "
                f"(exit code {speaking_process.exitcode})."
            )
        # Otherwise, keep waiting until readiness times out.
        if time.monotonic() >= deadline:
            speaking_process.terminate()
            speaking_process.join(timeout=TERMINATE_TIMEOUT_SECONDS)
            startup_error_receiver.close()
            raise ProcessStateError(
                "Speech process did not become ready within "
                f"{ready_timeout} seconds."
            )
        time.sleep(0.01)
    startup_error_receiver.close()
    return speaking_process


def start_speaking(
    speech_manager, voice_name, speech_rate, startup_error_sender=None
):
    """Initiate the speech process based on the speech manager's state."""
    try:
        speech_engine = win32com.client.Dispatch("SAPI.SpVoice")

        voices_collection = speech_engine.GetVoices()
        selected_sapi_voice_token = None
        if voice_name:
            for i in range(voices_collection.Count):
                voice_token = voices_collection.Item(i)
                if voice_name.lower() in voice_token.GetDescription().lower():
                    selected_sapi_voice_token = voice_token
                    break
        if selected_sapi_voice_token:
            speech_engine.Voice = selected_sapi_voice_token
        elif voices_collection.Count > 0:
            speech_engine.Voice = voices_collection.Item(0)

        if speech_rate:
            speech_engine.Rate = max(-10, min(10, speech_rate))

        speech_manager.set_ready(True)
    except Exception as e:
        if startup_error_sender is not None:
            startup_error_sender.send(f"{type(e).__name__}: {e}")
            startup_error_sender.close()
        raise
    if startup_error_sender is not None:
        startup_error_sender.close()
    while speech_manager.can_speak():
        text = speech_manager.pop_speech_text()
        if text:
            speech_engine.Speak(text)

        time.sleep(0.01)


def stop_speaking_process(
    base_manager,
    speech_manager,
    speaking_process,
    join_timeout=JOIN_TIMEOUT_SECONDS,
):
    """Stop the speaking process and shutdown the base manager."""
    if speech_manager.get_speech_text():
        time.sleep(0.01)

    try:
        speech_manager.set_can_speak(False)
        speaking_process.join(timeout=join_timeout)
        if speaking_process.is_alive():
            speaking_process.terminate()
            speaking_process.join(timeout=TERMINATE_TIMEOUT_SECONDS)
            raise ProcessStateError(
                "Speech process did not stop within "
                f"{join_timeout} seconds."
            )
    finally:
        base_manager.shutdown()
