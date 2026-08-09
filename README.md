# 🤖 Gemini AI Telegram Assistant & Coding Mentor

[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Telegram](https://img.shields.io/badge/Telegram-Bot_API-26A5E4?style=for-the-badge&logo=telegram&logoColor=white)](https://core.telegram.org/bots/api)
[![Gemini API](https://img.shields.io/badge/Google-Gemini_API-8E44AD?style=for-the-badge&logo=google&logoColor=white)](https://ai.google.dev/)

An intelligent, multimodal Telegram bot powered by the **Google Gemini API**. Built using `python-telegram-bot`, it functions as a contextual AI assistant, coding mentor, slide presenter, and multimedia generator.

---

## ✨ Key Features

* **💬 Contextual AI Chat:** Remembers user conversation history and acts as a programming mentor.
* **📊 Presentation Generator (`/slides`):** Generates structured PowerPoint (`.pptx`) decks on any topic using Gemini's structured JSON output and `python-pptx`.
* **🎨 Image Generation (`/image`):** Generates high-quality AI images using Imagen / Gemini vision models.
* **📸 Multimodal Capabilities:** Accepts photos, screenshots, and code snippets for visual analysis and debugging.
* **⏰ Smart Reminders (`/remind`):** Asynchronous task scheduling using `python-telegram-bot` JobQueue.
* **🌍 Bilingual Interface:** Built-in support for English and Russian.

---

## 🛠️ Tech Stack

* **Language:** Python 3.11+
* **Framework:** `python-telegram-bot` (v20+)
* **AI Provider:** `google-genai` (Google Gemini API)
* **Document Generation:** `python-pptx`
* **Environment Management:** `python-dotenv`

---

## 🚀 Quick Start

### 1. Clone the repository
```bash
git clone [https://github.com/erbulanbombi-ux/project-telegram-ai-bot.git](https://github.com/erbulanbombi-ux/project-telegram-ai-bot.git)
cd project-telegram-ai-bot