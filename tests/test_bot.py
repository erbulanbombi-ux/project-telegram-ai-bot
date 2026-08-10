import asyncio
import unittest
from unittest.mock import AsyncMock, patch

import bot


class FakeMessage:
    def __init__(self):
        self.text = "привет"
        self.caption = None
        self.photo = None
        self.document = None
        self.replies = []

    async def reply_text(self, text):
        self.replies.append(text)


class FakeChat:
    def __init__(self):
        self.actions = []

    async def send_action(self, action):
        self.actions.append(action)


class FakeUpdate:
    def __init__(self):
        self.effective_message = FakeMessage()
        self.effective_chat = FakeChat()


class FakeContext:
    def __init__(self):
        self.user_data = {}
        self.args = []


class BotChatTests(unittest.TestCase):
    def test_chat_returns_friendly_message_when_gemini_is_unavailable(self):
        async def run_test():
            update = FakeUpdate()
            context = FakeContext()

            with patch(
                "bot.ask_gemini_with_fallback",
                new=AsyncMock(side_effect=RuntimeError("429 RESOURCE_EXHAUSTED")),
            ):
                await bot.chat(update, context)

            self.assertIn(
                "Сейчас Gemini временно недоступен",
                update.effective_message.replies[-1],
            )

        asyncio.run(run_test())

    def test_retry_logic_treats_quota_error_as_transient(self):
        async def run_test():
            class FakeResponse:
                text = "Ответ после повтора"

            class FakeModels:
                async def generate_content(self, **kwargs):
                    return FakeResponse()

            class FakeClient:
                aio = type("aio", (), {"models": FakeModels()})

            original_client = bot.genai.Client
            bot.genai.Client = lambda api_key=None: FakeClient()
            try:
                response = await bot.ask_gemini_with_fallback(contents=["hi"])
            finally:
                bot.genai.Client = original_client

            self.assertEqual(response.text, "Ответ после повтора")

        asyncio.run(run_test())


if __name__ == "__main__":
    unittest.main()
