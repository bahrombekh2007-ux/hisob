# -*- coding: utf-8 -*-
"""
Render'da ishga tushirish nuqtasi.
Bitta process ichida:
  - Telegram bot (polling) alohida threadda ishlaydi
  - Flask web-server asosiy threadda ishlaydi va PORT'ni tinglaydi
Render "Start Command": python main.py
"""
import os
import threading
import time
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("main")

from app import app, init_db
from bot import start_bot


def run_bot_thread():
    while True:
        try:
            logger.info("Bot polling boshlanmoqda...")
            start_bot()
        except Exception:
            logger.exception("Bot xato bilan to'xtadi, 5 soniyadan keyin qayta urinamiz")
            time.sleep(5)


if __name__ == "__main__":
    init_db()

    bot_thread = threading.Thread(target=run_bot_thread, daemon=True)
    bot_thread.start()

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True, use_reloader=False)
