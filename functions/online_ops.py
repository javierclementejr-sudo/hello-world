"""Online operations for Jarvis assistant."""

import os
import json
import webbrowser
from urllib.parse import quote_plus

import requests

NEWS_API_KEY = os.getenv("NEWS_API_KEY", "")
OPENWEATHER_APP_ID = os.getenv("OPENWEATHER_APP_ID", "")
TMDB_API_KEY = os.getenv("TMDB_API_KEY", "")
EMAIL_ADDRESS = os.getenv("EMAIL", "")
EMAIL_PASSWORD = os.getenv("PASSWORD", "")
NEWS_COUNTRY = os.getenv("NEWS_COUNTRY", "es")
WHATSAPP_COUNTRY_CODE = os.getenv("WHATSAPP_COUNTRY_CODE", "")


def find_my_ip():
    response = requests.get("https://api64.ipify.org?format=json", timeout=10)
    response.raise_for_status()
    return response.json().get("ip", "desconocida")


def get_city_from_ip():
    response = requests.get("https://ipapi.co/json/", timeout=10)
    response.raise_for_status()
    data = response.json()
    return data.get("city", "Madrid")


def get_latest_news():
    if not NEWS_API_KEY:
        raise RuntimeError("NEWS_API_KEY no configurada")
    url = f"https://newsapi.org/v2/top-headlines?country={NEWS_COUNTRY}&apiKey={NEWS_API_KEY}"
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    articles = response.json().get("articles", [])
    return [article.get("title", "") for article in articles[:10] if article.get("title")]


def get_random_advice():
    response = requests.get("https://api.adviceslip.com/advice", timeout=10)
    response.raise_for_status()
    advice = response.json().get("slip", {}).get("advice", "")
    return f"Consejo: {advice}" if advice else "No pude obtener un consejo ahora."


def get_random_joke(tipo="normal"):
    from duckduckgo_search import DDGS
    query = f"chiste {tipo} en español"
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=1))
        if results:
            return results[0].get("body", "No encontre un chiste ahora.")
    except Exception:
        pass
    return "No pude encontrar un chiste ahora mismo."


def get_trending_movies():
    if not TMDB_API_KEY:
        raise RuntimeError("TMDB_API_KEY no configurada")
    url = f"https://api.themoviedb.org/3/trending/movie/week?api_key={TMDB_API_KEY}&language=es-ES"
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    movies = response.json().get("results", [])
    return [movie.get("title", "") for movie in movies[:10] if movie.get("title")]


def get_weather_report(city):
    if not OPENWEATHER_APP_ID:
        raise RuntimeError("OPENWEATHER_APP_ID no configurada")
    url = (
        f"https://api.openweathermap.org/data/2.5/weather"
        f"?q={quote_plus(city)}&appid={OPENWEATHER_APP_ID}&units=metric&lang=es"
    )
    response = requests.get(url, timeout=10)
    response.raise_for_status()
    data = response.json()
    weather = data.get("weather", [{}])[0].get("description", "desconocido")
    temp = data.get("main", {}).get("temp", "?")
    feels_like = data.get("main", {}).get("feels_like", "?")
    return weather, f"{temp}°C", f"{feels_like}°C"


def play_on_youtube(query):
    url = f"https://www.youtube.com/results?search_query={quote_plus(query)}"
    webbrowser.open(url)


def search_on_google(query):
    url = f"https://www.google.com/search?q={quote_plus(query)}"
    webbrowser.open(url)


def search_on_wikipedia(query):
    import wikipedia
    wikipedia.set_lang("es")
    return wikipedia.summary(query, sentences=3)


def send_email(to_address, subject, message):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    if not EMAIL_ADDRESS or not EMAIL_PASSWORD:
        return False, "Credenciales de correo no configuradas."
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = to_address
        msg["Subject"] = subject
        msg.attach(MIMEText(message, "plain"))

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.sendmail(EMAIL_ADDRESS, to_address, msg.as_string())
        server.quit()
        return True, None
    except Exception as e:
        return False, str(e)


def send_whatsapp_message(number, message):
    url = f"https://web.whatsapp.com/send?phone={number}&text={quote_plus(message)}"
    webbrowser.open(url)
