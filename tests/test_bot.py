import asyncio
import types as pytypes
import unittest

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
    def test_chat_returns_reply_without_crashing(self):
        async def run_test():
            class FakeResponse:
                text = "Привет! Я на связи."

            class FakeModels:
                async def generate_content(self, **kwargs):
                    return FakeResponse()

            class FakeClient:
                aio = pytypes.SimpleNamespace(models=FakeModels())

            original_client = bot.client
            bot.client = FakeClient()
            try:
                update = FakeUpdate()
                context = FakeContext()
                await bot.chat(update, context)
            finally:
                bot.client = original_client

            self.assertTrue(update.effective_message.replies)

        asyncio.run(run_test())


if __name__ == "__main__":
    unittest.main()
