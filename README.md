# J.A.R.V.I.S. — Stark Industries AI Assistant

A futuristic Iron Man-style voice assistant with an animated HUD interface, web search capabilities, voice interruption, and desktop automation.

## Features

- **Iron Man HUD Interface**: Animated arc reactor with rotating rings, pulsing core, scan lines, and HUD corner brackets. Color changes based on state (idle/listening/speaking).
- **Voice Control**: Wake word "Jarvis" activates listening. Continuous speech recognition with Google Speech API.
- **Voice Interruption**: Automatically stops speaking when it detects you're talking.
- **Web Search & Comparisons**: Can compare anything (products, countries, technologies) using DuckDuckGo and Wikipedia.
- **AI Brain**: Uses Groq (LLaMA 3.1) with Ollama fallback for intelligent responses.
- **Desktop Automation**: Open/close apps, control volume, media playback, screenshots, and more.
- **Cross-Platform**: Works on Windows and Linux with platform-specific fallbacks.

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in your API keys:
   ```bash
   cp .env.example .env
   ```
4. Run:
   ```bash
   python main.py
   ```

## Requirements

- Python 3.10+
- Microphone for voice control
- Internet connection for web search and AI responses

## Voice Commands

- **"Jarvis, qué hora es"** — Current time
- **"Jarvis, compara iPhone vs Samsung"** — Web comparison
- **"Jarvis, busca información sobre..."** — Web search
- **"Jarvis, abre Chrome"** — Open application
- **"Jarvis, volumen al 50"** — Volume control
- **"Jarvis, clima en Madrid"** — Weather report
- **"Jarvis, traduce 'hello' en español"** — Translation
- And many more...
