import json
from pathlib import Path
import sys
import tempfile
import unittest


SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from ppt_metrics import Metrics, _sizes, _token_estimate


class MetricsTests(unittest.TestCase):
    def metrics(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        return Metrics(Path(directory.name) / "metrics.jsonl")

    @staticmethod
    def success(result=None, **extra):
        response = {"ok": True, "result": result or {}}
        response.update(extra)
        return response

    def test_new_operation_labels_and_classification(self):
        metrics = self.metrics()
        for op in ("context", "outline", "review"):
            metrics.record(
                {"op": op, "deck": "deck-1", "slide_id": 4},
                self.success({"revision": f"{op}-revision"}),
                1,
            )
        metrics.record(
            {"op": "native_deck", "mode": "write"},
            self.success({"revision": "write-revision"}),
            1,
        )

        summary = metrics.summary()
        self.assertEqual(
            {"context": 1, "native_deck": 1, "outline": 1, "review": 1},
            {key: summary["operation_counts"][key] for key in ("context", "native_deck", "outline", "review")},
        )
        self.assertEqual(1, summary["mutation_requests"])

    def test_repeated_reads_and_renders_are_counted(self):
        metrics = self.metrics()

        deck_request = {"op": "deck", "request_id": "first", "name": "Quarterly Deck"}
        deck_response = self.success({"revision": "deck-revision"})
        metrics.record(deck_request, deck_response, 1)
        metrics.record({**deck_request, "request_id": "second"}, deck_response, 1)

        context_request = {"op": "context", "deck": "deck-1", "slide_id": 7}
        context_response = self.success({"active_slide_id": 7})
        metrics.record(context_request, context_response, 1)
        metrics.record(context_request, context_response, 1)

        outline_request = {"op": "outline", "deck": "deck-1"}
        outline_response = self.success({"revision": "outline-revision"})
        metrics.record(outline_request, outline_response, 1)
        metrics.record(outline_request, outline_response, 1)

        render_request = {"op": "render", "deck": "deck-1", "slide_id": 7, "width": 640}
        render_response = self.success({"image_hash": "image-1"})
        metrics.record(render_request, render_response, 1)
        metrics.record(render_request, render_response, 1)
        metrics.record(
            {**render_request, "width": 800},
            render_response,
            1,
        )

        review_request = {"op": "review", "deck": "deck-1", "slide_id": 7}
        review_response = self.success({"image_hash": "review-image-1"})
        metrics.record(review_request, review_response, 1)
        metrics.record(review_request, review_response, 1)

        summary = metrics.summary()
        flags = summary["flags"]
        self.assertEqual(1, flags["repeated_deck"])
        self.assertEqual(1, flags["repeated_context"])
        self.assertEqual(1, flags["repeated_outline"])
        self.assertEqual(2, flags["repeated_render"])
        self.assertGreater(summary["estimated_avoidable_text_tokens"], 0)

    def test_failed_calls_do_not_add_repeat_savings(self):
        metrics = self.metrics()
        request = {"op": "outline", "deck": "deck-1"}
        response = self.success({"revision": "outline-revision"})
        metrics.record(request, response, 1)
        before = metrics.summary()

        metrics.record(
            request,
            {
                "ok": False,
                "error": "outline failed",
                "result": {"revision": "outline-revision"},
            },
            1,
        )
        after = metrics.summary()
        self.assertEqual(before["estimated_avoidable_text_tokens"], after["estimated_avoidable_text_tokens"])
        self.assertEqual(0, after["flags"]["repeated_outline"])
        self.assertEqual(1, after["failed_requests"])

    def test_large_and_broad_reply_flags_do_not_create_savings(self):
        metrics = self.metrics()
        response = self.success(
            {"revision": "review-revision"},
            payload="reply-value-" + ("x" * 12_000),
        )
        metrics.record(
            {"op": "review", "deck": "deck-1", "slide_id": 3},
            response,
            1,
        )

        summary = metrics.summary()
        self.assertEqual(1, summary["flags"]["large_replies"])
        self.assertEqual(1, summary["flags"]["broad_reads"])
        self.assertEqual(0, summary["estimated_avoidable_text_tokens"])

    def test_log_contains_no_request_or_response_values(self):
        metrics = self.metrics()
        request_values = {
            "op": "context",
            "deck": "Deck Name That Must Stay Private",
            "path": r"C:\private\presentation.pptx",
            "text": "REQUEST_SECRET_VALUE",
        }
        response_values = {
            "revision": "REVISION_SECRET_VALUE",
            "name": "Response Presentation Name",
            "text": "RESPONSE_SECRET_VALUE",
        }
        metrics.record(request_values, self.success(response_values), 1)

        log_text = Path(metrics.summary()["log_path"]).read_text(encoding="utf-8")
        for value in (
            request_values["deck"],
            request_values["path"],
            request_values["text"],
            response_values["revision"],
            response_values["name"],
            response_values["text"],
        ):
            self.assertNotIn(value, log_text)
        self.assertIn('"op":"context"', log_text)

    def test_avoidable_count_matches_only_the_repeated_call(self):
        metrics = self.metrics()
        request = {"op": "deck", "deck": "deck-1"}
        response = self.success({"revision": "revision-1"})
        metrics.record(request, response, 1)
        metrics.record(request, response, 1)

        request_tokens = _token_estimate(_sizes(request)[0])
        response_tokens = _token_estimate(_sizes(response)[0])
        self.assertEqual(
            request_tokens + response_tokens,
            metrics.summary()["estimated_avoidable_text_tokens"],
        )


if __name__ == "__main__":
    unittest.main()
