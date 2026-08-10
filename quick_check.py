import bot
print('is_unavail:', bot.is_gemini_unavailable_error(RuntimeError('429 RESOURCE_EXHAUSTED')))
print('format:', bot.format_gemini_error_message(RuntimeError('429 RESOURCE_EXHAUSTED')))
