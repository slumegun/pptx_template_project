from threading import Barrier

from engine.models import SlideContent
from engine.pipeline import _run_workers


def test_workers_use_the_assigned_model_concurrently_and_preserve_slide_order(monkeypatch):
    first = object()
    barrier = Barrier(2)
    observed = []
    slides = [SlideContent("slide_1", "One", ["a"], ["f1"]), SlideContent("slide_2", "Two", ["b"], ["f2"])]

    def refine(client, slide, facts, outline, deadline, brief):
        observed.append(client)
        barrier.wait(timeout=3)
        return slide

    monkeypatch.setenv("MODEL_ENABLE_WORKERS", "0")
    monkeypatch.setenv("MODEL_MAX_PARALLEL", "2")
    monkeypatch.setattr("engine.pipeline._refine_slide", refine)
    assert _run_workers(first, slides, []) == slides
    assert set(observed) == {first}
