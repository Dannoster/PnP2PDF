import threading

from pnp2pdf.errors import OperationCancelled
from pnp2pdf.ui.job_runner import BuildJob


def _recorder(events):
    return lambda name, payload: events.append((name, payload))


def test_success_posts_done():
    events = []

    def build(options, progress=None, cancelled=None):
        return "/tmp/out"

    job = BuildJob(object(), _recorder(events), build=build)
    assert job.start() is True
    job.join(5)

    assert ("done", "/tmp/out") in events
    assert job.path == "/tmp/out"


def test_failure_posts_error():
    events = []

    def build(options, progress=None, cancelled=None):
        raise RuntimeError("boom")

    job = BuildJob(object(), _recorder(events), build=build)
    job.start()
    job.join(5)

    assert ("error", "boom") in events


def test_cancellation_posts_cancelled():
    events = []

    def build(options, progress=None, cancelled=None):
        while not cancelled():
            pass
        raise OperationCancelled()

    job = BuildJob(object(), _recorder(events), build=build)
    job.start()
    job.cancel()
    job.join(5)

    assert ("cancelled", None) in events


def test_build_runs_off_main_thread():
    events = []

    def build(options, progress=None, cancelled=None):
        assert threading.current_thread() is not threading.main_thread()
        return "/tmp/out"

    job = BuildJob(object(), _recorder(events), build=build)
    job.start()
    job.join(5)

    assert ("done", "/tmp/out") in events


def test_start_twice_returns_false():
    gate = threading.Event()

    def build(options, progress=None, cancelled=None):
        gate.wait(5)
        return "/tmp/out"

    job = BuildJob(object(), _recorder([]), build=build)
    assert job.start() is True
    assert job.start() is False
    gate.set()
    job.join(5)
