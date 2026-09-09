import importlib.util
import json
import os
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('guard', Path(__file__).parents[1] / 'scripts/guard.py')
guard = importlib.util.module_from_spec(spec)
spec.loader.exec_module(guard)
NOW = datetime(2026, 9, 6, 12, tzinfo=timezone.utc)


class GuardTests(unittest.TestCase):
    def check(self, platform, tokens=50000, at='2026-09-06T10:00:00Z', hook='UserPromptSubmit', use_defaults=False,
              prompt=''):
        if platform == 'claude':
            row = {'type': 'assistant', 'timestamp': at, 'message': {'usage': {
                'input_tokens': 1000, 'cache_read_input_tokens': tokens - 1000},
                'content': [{'type': 'text', 'text': 'Task state'}]}}
        else:
            row = {'type': 'event_msg', 'timestamp': at, 'payload': {'type': 'token_count',
                   'info': {'last_token_usage': {'input_tokens': tokens, 'cached_input_tokens': tokens},
                            'total_token_usage': {'input_tokens': 999999}}}}
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'session.jsonl'
            path.write_text('malformed\n[]\n' + json.dumps(row) + '\n')
            with patch.dict(os.environ, {} if use_defaults else {
                    'TOKENSAVER_CLAUDE_TTL_SECONDS': '3600',
                    'TOKENSAVER_CODEX_TTL_SECONDS': '3600',
                    'TOKENSAVER_THRESHOLD_TOKENS': '50000'}, clear=True):
                return guard.evaluate({'hook_event_name': hook, 'source': 'resume', 'prompt': prompt,
                                       'transcript_path': str(path)}, platform, NOW)

    def test_stale_threshold_both_platforms(self):
        for platform in ('claude', 'codex'):
            self.assertEqual(self.check(platform)['tokens'], 50000)

    def test_continue_prompt_overrides_block(self):
        for platform in ('claude', 'codex'):
            self.assertIsNone(self.check(platform, prompt='  Continue  '))
            self.assertIsNotNone(self.check(platform, prompt='continue the refactor'))

    def test_small_and_fresh_allowed(self):
        for platform in ('claude', 'codex'):
            self.assertIsNone(self.check(platform, 49999))
            self.assertIsNone(self.check(platform, at='2026-09-06T11:59:00Z'))

    def test_subscription_main_thread_default_boundaries(self):
        self.assertIsNone(self.check('codex', at='2026-09-06T11:40:00Z', use_defaults=True))
        for platform, boundary, expired in (
                ('codex', '2026-09-06T11:30:00Z', '2026-09-06T11:29:59Z'),
                ('claude', '2026-09-06T11:00:00Z', '2026-09-06T10:59:59Z')):
            with self.subTest(platform=platform):
                self.assertIsNone(self.check(platform, at=boundary, use_defaults=True))
                self.assertIsNotNone(self.check(platform, at=expired, use_defaults=True))
                self.assertIsNone(self.check(platform, tokens=49999, at=expired, use_defaults=True))

    def test_ttl_override_is_respected(self):
        self.assertIsNone(self.check('codex', at='2026-09-06T11:20:00Z'))
        self.assertIsNotNone(self.check('codex', at='2026-09-06T11:20:00Z', use_defaults=True))

    def test_session_start_platform_behavior(self):
        self.assertIsNone(self.check('claude', hook='SessionStart'))
        self.assertIsNotNone(self.check('codex', hook='SessionStart'))

    def test_unknown_timestamp_allowed(self):
        self.assertIsNone(self.check('codex', at='invalid'))

    def test_missing_path_allowed(self):
        self.assertIsNone(guard.evaluate({'hook_event_name': 'UserPromptSubmit'}, 'claude'))

    def test_private_bounded_handoff(self):
        data = self.check('claude')
        path = guard.save_handoff({'session_id': '../../escape'}, data)
        try:
            self.assertEqual(path.stat().st_mode & 0o777, 0o600)
            self.assertEqual(path.parent.stat().st_mode & 0o777, 0o700)
            self.assertEqual(json.loads(path.read_text())['recent_exchanges'][0]['text'], 'Task state')
        finally:
            path.unlink()
            path.parent.rmdir()


if __name__ == '__main__':
    unittest.main()
