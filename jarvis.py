import difflib
import glob
import json
import math
import os
import platform
import re
import shutil
import subprocess
import threading
import time
import tkinter as tk
import webbrowser
from urllib.parse import quote_plus
from pathlib import Path
from datetime import datetime, timedelta

import customtkinter as ctk
import pyautogui
import pyttsx3
import speech_recognition as sr
import wikipedia
from duckduckgo_search import DDGS

try:
    from groq import Groq
except ImportError:
    Groq = None

try:
    import ollama
except ImportError:
    ollama = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import pytesseract

    if platform.system() == "Windows":
        pytesseract.pytesseract.tesseract_cmd = (
            r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        )
except Exception:
    pytesseract = None

IS_WINDOWS = platform.system() == "Windows"

if IS_WINDOWS:
    try:
        from ctypes import (
            POINTER,
            Structure,
            WINFUNCTYPE,
            byref,
            c_float,
            windll,
        )
        from ctypes import wintypes

        HAS_WIN_AUDIO = True
    except Exception:
        HAS_WIN_AUDIO = False
else:
    HAS_WIN_AUDIO = False

from functions.online_ops import (
    find_my_ip,
    get_city_from_ip,
    get_latest_news,
    get_random_advice,
    get_random_joke,
    get_trending_movies,
    get_weather_report,
    play_on_youtube,
    search_on_google,
    search_on_wikipedia,
    send_email,
    send_whatsapp_message,
)

USERNAME = "Señor Clemente"
BOTNAME = "Jarvis"


def speak(texto):
    try:
        engine = pyttsx3.init()
        engine.say(texto)
        engine.runAndWait()
    except Exception:
        pass


def greet_user():
    hour = datetime.now().hour
    if 6 <= hour < 12:
        speak(f"Buenos días {USERNAME}")
    elif 12 <= hour < 20:
        speak(f"Buenas tardes {USERNAME}")
    else:
        speak(f"Buenas noches {USERNAME}")
    speak(f"Yo soy {BOTNAME}. ¿Cómo puedo asistirle, señor?")


WAKE_WORDS = (
    "jarvis",
    "jarbis",
    "yarvis",
    "yarbis",
    "jarbi",
    "yarbi",
    "harvest",
)
CARPETA_CAPTURAS = os.path.join(os.path.dirname(__file__), "capturas")

# === CONFIGURACION ===
API_KEY = os.getenv("GROQ_API_KEY", "").strip()
client = None
ARCHIVO_MEMORIA = os.path.join(os.path.dirname(__file__), "memoria.json")
ARCHIVO_AGENDA = os.path.join(os.path.dirname(__file__), "agenda.json")
wikipedia.set_lang("es")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
pyautogui.FAILSAFE = True

APLICACIONES_CONOCIDAS = {
    "spotify": {
        "launch": "spotify",
        "processes": ["Spotify.exe", "spotify"],
        "window_titles": ["spotify"],
    },
    "discord": {
        "launch": "discord",
        "processes": ["Discord.exe", "discord"],
        "window_titles": ["discord"],
    },
    "chrome": {
        "launch": "google-chrome" if not IS_WINDOWS else "chrome",
        "processes": ["chrome.exe", "chrome", "google-chrome"],
        "window_titles": ["google chrome", "chrome"],
    },
    "google chrome": {
        "launch": "google-chrome" if not IS_WINDOWS else "chrome",
        "processes": ["chrome.exe", "chrome", "google-chrome"],
        "window_titles": ["google chrome", "chrome"],
    },
    "firefox": {
        "launch": "firefox",
        "processes": ["firefox.exe", "firefox"],
        "window_titles": ["firefox", "mozilla firefox"],
    },
    "edge": {
        "launch": "msedge" if IS_WINDOWS else "microsoft-edge",
        "processes": ["msedge.exe", "msedge"],
        "window_titles": ["microsoft edge", "edge"],
    },
    "explorador": {
        "launch": "explorer" if IS_WINDOWS else "nautilus",
        "processes": ["explorer.exe", "nautilus"],
        "window_titles": ["explorador de archivos", "file explorer", "files"],
    },
    "explorador de archivos": {
        "launch": "explorer" if IS_WINDOWS else "nautilus",
        "processes": ["explorer.exe", "nautilus"],
        "window_titles": ["explorador de archivos", "file explorer", "files"],
    },
    "notepad": {
        "launch": "notepad" if IS_WINDOWS else "gedit",
        "processes": ["notepad.exe", "gedit"],
        "window_titles": ["bloc de notas", "notepad", "gedit"],
    },
    "bloc de notas": {
        "launch": "notepad" if IS_WINDOWS else "gedit",
        "processes": ["notepad.exe", "gedit"],
        "window_titles": ["bloc de notas", "notepad", "gedit"],
    },
    "vs code": {
        "launch": "code",
        "processes": ["Code.exe", "code"],
        "window_titles": ["visual studio code", "vs code"],
    },
    "visual studio code": {
        "launch": "code",
        "processes": ["Code.exe", "code"],
        "window_titles": ["visual studio code", "vs code"],
    },
    "telegram": {
        "launch": "telegram" if IS_WINDOWS else "telegram-desktop",
        "processes": ["Telegram.exe", "telegram-desktop"],
        "window_titles": ["telegram"],
    },
    "intellij": {
        "launch": "idea64" if IS_WINDOWS else "idea",
        "processes": ["idea64.exe", "idea"],
        "window_titles": ["intellij", "idea"],
    },
    "intellij idea": {
        "launch": "idea64" if IS_WINDOWS else "idea",
        "processes": ["idea64.exe", "idea"],
        "window_titles": ["intellij", "idea"],
    },
    "mysql workbench": {
        "launch": "MySQLWorkbench",
        "processes": ["MySQLWorkbench.exe", "mysql-workbench"],
        "window_titles": ["mysql workbench"],
    },
    "dbeaver": {
        "launch": "dbeaver",
        "processes": ["dbeaver.exe", "dbeaver"],
        "window_titles": ["dbeaver"],
    },
    "windsurf": {
        "launch": "windsurf",
        "processes": ["windsurf.exe", "windsurf"],
        "window_titles": ["windsurf"],
    },
    "opera gx": {
        "launch": "opera",
        "processes": ["opera.exe", "opera"],
        "window_titles": ["opera", "opera gx"],
    },
    "opera": {
        "launch": "opera",
        "processes": ["opera.exe", "opera"],
        "window_titles": ["opera", "opera gx"],
    },
}

# Windows media key constants
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF
VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_MEDIA_STOP = 0xB2
VK_MEDIA_PLAY_PAUSE = 0xB3
KEYEVENTF_KEYUP = 0x0002

if HAS_WIN_AUDIO:
    CLSCTX_ALL = 23
    COINIT_APARTMENTTHREADED = 0x2
    HRESULT = wintypes.LONG

    class GUID(Structure):
        _fields_ = [
            ("Data1", wintypes.DWORD),
            ("Data2", wintypes.WORD),
            ("Data3", wintypes.WORD),
            ("Data4", wintypes.BYTE * 8),
        ]

        def __init__(self, guid_str):
            super().__init__()
            guid_str = guid_str.strip("{}")
            data1, data2, data3, data4a, data4b = guid_str.split("-")
            self.Data1 = int(data1, 16)
            self.Data2 = int(data2, 16)
            self.Data3 = int(data3, 16)
            bytes_data = bytes.fromhex(data4a + data4b)
            self.Data4[:] = bytes_data

    class PROPERTYKEY(Structure):
        _fields_ = [("fmtid", GUID), ("pid", wintypes.DWORD)]

    class IUnknown:
        _iid_ = GUID("{00000000-0000-0000-C000-000000000046}")

    class IMMDeviceEnumerator(IUnknown):
        _iid_ = GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")

    class IMMDevice(IUnknown):
        _iid_ = GUID("{D666063F-1587-4E43-81F1-B948E807363F}")

    class IAudioEndpointVolume(IUnknown):
        _iid_ = GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}")

    class PROPVARIANT_UNION(Structure):
        _fields_ = [("pwszVal", wintypes.LPWSTR)]

    class PROPVARIANT(Structure):
        _anonymous_ = ("u",)
        _fields_ = [
            ("vt", wintypes.USHORT),
            ("wReserved1", wintypes.USHORT),
            ("wReserved2", wintypes.USHORT),
            ("wReserved3", wintypes.USHORT),
            ("u", PROPVARIANT_UNION),
        ]

    LP_GUID = POINTER(GUID)

    class IUnknownVtbl(Structure):
        _fields_ = [
            ("QueryInterface", wintypes.LPVOID),
            ("AddRef", wintypes.LPVOID),
            ("Release", wintypes.LPVOID),
        ]

    class IMMDeviceEnumeratorVtbl(Structure):
        _fields_ = [
            ("QueryInterface", wintypes.LPVOID),
            ("AddRef", wintypes.LPVOID),
            ("Release", wintypes.LPVOID),
            ("EnumAudioEndpoints", wintypes.LPVOID),
            (
                "GetDefaultAudioEndpoint",
                WINFUNCTYPE(
                    HRESULT,
                    wintypes.LPVOID,
                    wintypes.DWORD,
                    wintypes.DWORD,
                    wintypes.LPVOID,
                ),
            ),
            ("GetDevice", wintypes.LPVOID),
            ("RegisterEndpointNotificationCallback", wintypes.LPVOID),
            ("UnregisterEndpointNotificationCallback", wintypes.LPVOID),
        ]

    class IMMDeviceVtbl(Structure):
        _fields_ = [
            ("QueryInterface", wintypes.LPVOID),
            ("AddRef", wintypes.LPVOID),
            ("Release", wintypes.LPVOID),
            (
                "Activate",
                WINFUNCTYPE(
                    HRESULT,
                    wintypes.LPVOID,
                    LP_GUID,
                    wintypes.DWORD,
                    wintypes.LPVOID,
                    wintypes.LPVOID,
                ),
            ),
            ("OpenPropertyStore", wintypes.LPVOID),
            ("GetId", wintypes.LPVOID),
            ("GetState", wintypes.LPVOID),
        ]

    class IAudioEndpointVolumeVtbl(Structure):
        _fields_ = [
            ("QueryInterface", wintypes.LPVOID),
            ("AddRef", wintypes.LPVOID),
            ("Release", wintypes.LPVOID),
            ("RegisterControlChangeNotify", wintypes.LPVOID),
            ("UnregisterControlChangeNotify", wintypes.LPVOID),
            ("GetChannelCount", wintypes.LPVOID),
            ("SetMasterVolumeLevel", wintypes.LPVOID),
            (
                "SetMasterVolumeLevelScalar",
                WINFUNCTYPE(HRESULT, wintypes.LPVOID, c_float, wintypes.LPVOID),
            ),
            ("GetMasterVolumeLevel", wintypes.LPVOID),
            (
                "GetMasterVolumeLevelScalar",
                WINFUNCTYPE(HRESULT, wintypes.LPVOID, wintypes.LPVOID),
            ),
            ("SetChannelVolumeLevel", wintypes.LPVOID),
            ("SetChannelVolumeLevelScalar", wintypes.LPVOID),
            ("GetChannelVolumeLevel", wintypes.LPVOID),
            ("GetChannelVolumeLevelScalar", wintypes.LPVOID),
            (
                "SetMute",
                WINFUNCTYPE(
                    HRESULT, wintypes.LPVOID, wintypes.BOOL, wintypes.LPVOID
                ),
            ),
            ("GetMute", wintypes.LPVOID),
            ("GetVolumeStepInfo", wintypes.LPVOID),
            ("VolumeStepUp", wintypes.LPVOID),
            ("VolumeStepDown", wintypes.LPVOID),
            ("QueryHardwareSupport", wintypes.LPVOID),
            ("GetVolumeRange", wintypes.LPVOID),
        ]

    IMMDeviceEnumerator._fields_ = [
        ("lpVtbl", POINTER(IMMDeviceEnumeratorVtbl))
    ]
    IMMDevice._fields_ = [("lpVtbl", POINTER(IMMDeviceVtbl))]
    IAudioEndpointVolume._fields_ = [
        ("lpVtbl", POINTER(IAudioEndpointVolumeVtbl))
    ]


# === FUNCIONES DE APOYO ===
def cargar_memoria():
    if os.path.exists(ARCHIVO_MEMORIA):
        try:
            with open(ARCHIVO_MEMORIA, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return [
        {
            "role": "system",
            "content": (
                "Eres Jarvis, un asistente inteligente estilo Iron Man con acceso al PC, "
                "archivos, pantalla, automatizacion de escritorio y busqueda web avanzada. "
                "Puedes hacer comparativas entre productos, tecnologias, paises, "
                "deportistas, celebridades, o cualquier tema usando informacion de internet. "
                "IMPORTANTE: Todas tus respuestas se leen en voz alta. Sé conciso y directo. "
                "Evita listas muy largas, formatos complejos o texto excesivo. "
                "Si la información es extensa, da un resumen claro en máximo 3-4 frases. "
                "Siempre responde en español y siempre dirígete al usuario como 'señor'. "
                "Cuando el usuario pregunte sobre una persona famosa, celebridad, deportista, "
                "actor, político, etc., da información PRECISA y VERIFICADA: nombre completo, "
                "nacionalidad, profesión, logros principales y datos relevantes. "
                "Cuando el usuario pida traducir, da SOLO la traducción exacta. "
                "Nunca inventes datos. Si no estás seguro, dilo."
            ),
        }
    ]


def obtener_cliente_groq():
    global client
    if client is not None:
        return client
    if not API_KEY:
        return None
    if Groq is None:
        return None
    try:
        client = Groq(api_key=API_KEY)
    except Exception:
        client = None
    return client


def guardar_memoria(historial):
    with open(ARCHIVO_MEMORIA, "w", encoding="utf-8") as f:
        json.dump(historial, f, ensure_ascii=False, indent=4)


def cargar_agenda():
    if os.path.exists(ARCHIVO_AGENDA):
        try:
            with open(ARCHIVO_AGENDA, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    data.setdefault("items", [])
                    return data
        except Exception:
            pass
    return {"items": []}


def guardar_agenda(data):
    with open(ARCHIVO_AGENDA, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _agregar_item_agenda(tipo, mensaje, when_dt, metadata=None):
    agenda = cargar_agenda()
    items = agenda.get("items", [])
    item_id = f"{tipo}_{int(time.time() * 1000)}"
    items.append(
        {
            "id": item_id,
            "tipo": tipo,
            "mensaje": mensaje,
            "when_iso": when_dt.isoformat(),
            "notificado": False,
            "metadata": metadata or {},
        }
    )
    agenda["items"] = items
    guardar_agenda(agenda)
    return item_id


def _extraer_hora(texto):
    texto = texto.lower()
    marcas = [".", "h", ","]
    for marca in marcas:
        texto = texto.replace(marca, ":")
    tokens = texto.split()
    for token in tokens:
        if ":" in token:
            partes = token.split(":")
            if len(partes) >= 2 and partes[0].isdigit() and partes[1].isdigit():
                hora = int(partes[0])
                minuto = int(partes[1][:2])
                if 0 <= hora <= 23 and 0 <= minuto <= 59:
                    return hora, minuto
    return None, None


def _extraer_minutos_temporizador(texto):
    texto = texto.lower()
    numero = extraer_numero(texto)
    if numero is None:
        return None
    if "hora" in texto:
        return numero * 60
    return numero


def crear_temporizador(comando_original):
    minutos = _extraer_minutos_temporizador(comando_original)
    if minutos is None:
        return "No entendí la duración. Ejemplo: pon un temporizador de 10 minutos."
    motivo = "temporizador"
    if " para " in comando_original.lower():
        idx = comando_original.lower().find(" para ")
        motivo = comando_original[idx + len(" para ") :].strip() or motivo
    when_dt = datetime.now() + timedelta(minutes=minutos)
    _agregar_item_agenda(
        "temporizador",
        f"Temporizador: {motivo}",
        when_dt,
        {"minutos": minutos},
    )
    return f"Temporizador listo para {minutos} minutos. Avisaré cuando termine."


def crear_alarma(comando_original):
    hora, minuto = _extraer_hora(comando_original)
    if hora is None:
        return "No entendí la hora de la alarma. Ejemplo: despiértame a las 7:00."
    ahora = datetime.now()
    when_dt = ahora.replace(hour=hora, minute=minuto, second=0, microsecond=0)
    if when_dt <= ahora:
        when_dt += timedelta(days=1)
    _agregar_item_agenda("alarma", "Alarma programada", when_dt)
    return f"Alarma creada para las {hora:02d}:{minuto:02d}."


def crear_evento_calendario(comando_original):
    texto = comando_original.strip()
    titulo = "Evento"
    if "'" in texto:
        partes = texto.split("'")
        if len(partes) >= 3:
            titulo = partes[1].strip() or titulo

    dias = {
        "lunes": 0,
        "martes": 1,
        "miercoles": 2,
        "miércoles": 2,
        "jueves": 3,
        "viernes": 4,
        "sabado": 5,
        "sábado": 5,
        "domingo": 6,
    }
    target_day = None
    texto_lower = texto.lower()
    for nombre, idx in dias.items():
        if nombre in texto_lower:
            target_day = idx
            break

    hora, minuto = _extraer_hora(texto)
    if hora is None:
        hora, minuto = 10, 0

    now = datetime.now()
    if target_day is None:
        when_dt = now.replace(hour=hora, minute=minuto, second=0, microsecond=0)
        if when_dt <= now:
            when_dt += timedelta(days=1)
    else:
        delta = (target_day - now.weekday()) % 7
        if delta == 0:
            delta = 7
        when_dt = (now + timedelta(days=delta)).replace(
            hour=hora, minute=minuto, second=0, microsecond=0
        )

    _agregar_item_agenda("evento", titulo, when_dt)
    return f"Evento '{titulo}' añadido para {when_dt.strftime('%A %H:%M')}."


def crear_recordatorio(comando_original):
    texto = comando_original.strip()
    when_dt = datetime.now() + timedelta(minutes=30)
    trigger = "recordatorio general"
    if "cuando llegue" in texto.lower():
        trigger = "cuando llegue al destino"
    _agregar_item_agenda(
        "recordatorio", texto, when_dt, {"trigger": trigger}
    )
    return "Recordatorio guardado. Te avisaré en 30 minutos."


def obtener_items_vencidos_agenda():
    agenda = cargar_agenda()
    items = agenda.get("items", [])
    now = datetime.now()
    vencidos = []
    updated = False
    for item in items:
        if item.get("notificado"):
            continue
        when_iso = item.get("when_iso")
        if not when_iso:
            continue
        try:
            when_dt = datetime.fromisoformat(when_iso)
        except Exception:
            continue
        if when_dt <= now:
            item["notificado"] = True
            vencidos.append(item)
            updated = True
    if updated:
        agenda["items"] = items
        guardar_agenda(agenda)
    return vencidos


def convertir_unidades(comando_original):
    t = comando_original.lower()
    nums = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", t)]
    if ("celsius" in t or "°c" in t) and ("fahrenheit" in t or "°f" in t):
        if not nums:
            return "Indícame los grados. Ejemplo: 180 celsius a fahrenheit."
        c = nums[0]
        f = (c * 9 / 5) + 32
        return f"{c}°C son {f:.2f}°F."
    if ("fahrenheit" in t or "°f" in t) and ("celsius" in t or "°c" in t):
        if not nums:
            return "Indícame los grados. Ejemplo: 80 fahrenheit a celsius."
        f = nums[0]
        c = (f - 32) * 5 / 9
        return f"{f}°F son {c:.2f}°C."
    if ("dolar" in t or "dólar" in t or "usd" in t) and (
        "euro" in t or "eur" in t
    ):
        if not nums:
            return "Indícame cantidad. Ejemplo: cuantos euros son 50 dolares."
        cantidad = nums[0]
        try:
            import requests

            r = requests.get(
                "https://open.er-api.com/v6/latest/USD", timeout=10
            )
            r.raise_for_status()
            eur = r.json().get("rates", {}).get("EUR")
            if not eur:
                raise RuntimeError("sin tasa EUR")
            return f"{cantidad} USD son aproximadamente {cantidad * float(eur):.2f} EUR."
        except Exception:
            return "No pude consultar divisa en vivo ahora mismo."
    return None


IDIOMAS_MAP = {
    "inglés": "en", "ingles": "en", "english": "en",
    "francés": "fr", "frances": "fr", "french": "fr",
    "alemán": "de", "aleman": "de", "german": "de",
    "italiano": "it", "italian": "it",
    "portugués": "pt", "portugues": "pt", "portuguese": "pt",
    "japonés": "ja", "japones": "ja", "japanese": "ja",
    "chino": "zh", "chinese": "zh", "mandarín": "zh", "mandarin": "zh",
    "coreano": "ko", "korean": "ko",
    "ruso": "ru", "russian": "ru",
    "árabe": "ar", "arabe": "ar", "arabic": "ar",
    "hindi": "hi",
    "turco": "tr", "turkish": "tr",
    "polaco": "pl", "polish": "pl",
    "holandés": "nl", "holandes": "nl", "dutch": "nl",
    "sueco": "sv", "swedish": "sv",
    "catalán": "ca", "catalan": "ca",
    "griego": "el", "greek": "el",
    "hebreo": "he", "hebrew": "he",
    "tailandés": "th", "tailandes": "th",
    "vietnamita": "vi", "vietnamese": "vi",
    "rumano": "ro", "romanian": "ro",
    "checo": "cs", "czech": "cs",
    "húngaro": "hu", "hungaro": "hu", "hungarian": "hu",
    "español": "es", "espanol": "es", "spanish": "es",
    "latín": "la", "latin": "la",
}


def traducir_texto(comando_original):
    texto = comando_original.strip()
    frase = ""
    idioma = ""
    m = re.search(r"[\"']([^\"']+)[\"']", texto)
    if m:
        frase = m.group(1).strip()
        resto = texto[m.end() :].lower().replace("?", "").strip()
        if " en " in resto:
            idioma = resto.split(" en ", 1)[1].strip()
        elif " al " in resto:
            idioma = resto.split(" al ", 1)[1].strip()
        elif " a " in resto:
            idioma = resto.split(" a ", 1)[1].strip()
    if not frase and "traduce " in texto.lower():
        resto = texto.lower().split("traduce ", 1)[1]
        if " al " in resto:
            frase, idioma = resto.split(" al ", 1)
        elif " en " in resto:
            frase, idioma = resto.split(" en ", 1)
        elif " a " in resto:
            frase, idioma = resto.split(" a ", 1)
        frase = frase.strip(" .,:;")
        idioma = idioma.strip(" .,:;?")
    if not frase and "como se dice" in texto.lower():
        resto = texto.lower().split("como se dice", 1)[1].strip()
        if " en " in resto:
            frase, idioma = resto.split(" en ", 1)
        frase = frase.strip(" .,:;?")
        idioma = idioma.strip(" .,:;?")
    if not frase and "cómo se dice" in texto.lower():
        resto = texto.lower().split("cómo se dice", 1)[1].strip()
        if " en " in resto:
            frase, idioma = resto.split(" en ", 1)
        frase = frase.strip(" .,:;?")
        idioma = idioma.strip(" .,:;?")
    if not frase or not idioma:
        return "Usa: traduce 'tengo hambre' en japonés. O: cómo se dice hola en inglés."

    idioma_clean = idioma.strip().lower()
    idioma_code = IDIOMAS_MAP.get(idioma_clean, "")

    prompt = (
        f"Eres un traductor profesional. Traduce la siguiente frase al {idioma}. "
        f"Responde SOLO con la traducción exacta, sin explicaciones ni notas adicionales.\n\n"
        f"Frase: {frase}"
    )
    try:
        cliente_groq = obtener_cliente_groq()
        if cliente_groq:
            completion = cliente_groq.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": "Eres un traductor profesional. Solo responde con la traducción exacta."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.1,
            )
            respuesta = (completion.choices[0].message.content or "").strip()
            if respuesta:
                return f"'{frase}' en {idioma}: {respuesta}"
    except Exception:
        pass
    if ollama is not None:
        try:
            response = ollama.chat(
                model="llama3.2:1b",
                messages=[{"role": "user", "content": prompt}],
            )
            respuesta = (response["message"]["content"] or "").strip()
            if respuesta:
                return f"'{frase}' en {idioma}: {respuesta}"
        except Exception:
            pass
    tl = idioma_code or "auto"
    webbrowser.open(
        f"https://translate.google.com/?sl=auto&tl={tl}&text={quote_plus(frase)}&op=translate"
    )
    return f"No pude traducir ahora; abrí Google Translate para '{frase}' en {idioma}."


def calcular_rapido(comando_original):
    t = comando_original.lower()
    if "%" in t and "de" in t:
        nums = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", t)]
        if len(nums) >= 2:
            return f"El {nums[0]}% de {nums[1]} es {nums[1] * nums[0] / 100:.2f}."
    if t.startswith("calcula "):
        expr = t[len("calcula ") :].replace(",", ".")
        if re.fullmatch(r"[\d\.\+\-\*\/\(\)\s]+", expr):
            try:
                value = eval(expr, {"__builtins__": {}}, {})
                return f"Resultado: {value}"
            except Exception:
                return "No pude calcular esa expresión."
    return None


def obtener_resumen_web(consulta, max_results=3):
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=max_results))
        if not resultados:
            return "No encontré resultados ahora mismo."
        lineas = []
        for r in resultados[:max_results]:
            titulo = (r.get("title") or "").strip()
            body = (r.get("body") or "").strip()
            href = (r.get("href") or "").strip()
            linea = f"- {titulo}: {body}" if titulo else f"- {body}"
            if href:
                linea += f" [{href}]"
            lineas.append(linea.strip())
        return "\n".join(lineas)
    except Exception:
        return "No pude consultar esa información en tiempo real."


def realizar_comparativa_web(tema1, tema2, aspecto=""):
    """Searches the web for comparison data between two topics."""
    queries = [
        f"{tema1} vs {tema2} comparativa {aspecto}".strip(),
        f"{tema1} características especificaciones {aspecto}".strip(),
        f"{tema2} características especificaciones {aspecto}".strip(),
    ]
    resultados_combinados = []
    for query in queries:
        try:
            with DDGS() as ddgs:
                resultados = list(ddgs.text(query, max_results=3))
            for r in resultados:
                titulo = (r.get("title") or "").strip()
                body = (r.get("body") or "").strip()
                if body:
                    resultados_combinados.append(f"- {titulo}: {body}")
        except Exception:
            continue
    if not resultados_combinados:
        return "No encontré datos para la comparativa."
    return "\n".join(resultados_combinados[:8])


def detectar_comparativa(comando):
    """Detects if the user wants a comparison and extracts the items."""
    comando_lower = comando.lower()
    patterns = [
        r"compar[ae]\s+(.+?)\s+(?:con|vs|versus|y|contra)\s+(.+)",
        r"(?:que|qué|cual|cuál)\s+es\s+mejor\s+(.+?)\s+(?:o|vs)\s+(.+)",
        r"diferencias?\s+entre\s+(.+?)\s+y\s+(.+)",
        r"(.+?)\s+vs\.?\s+(.+)",
        r"(.+?)\s+versus\s+(.+)",
        r"(.+?)\s+o\s+(.+?)[\s,]*(?:cual|cuál|que|qué)\s+(?:es|sale)\s+mejor",
    ]
    for pattern in patterns:
        match = re.search(pattern, comando_lower)
        if match:
            return match.group(1).strip(" ?,"), match.group(2).strip(" ?.,")
    return None, None


def buscar_informacion_web(consulta, max_results=5):
    """Enhanced web search that fetches from multiple sources."""
    resultados_totales = []
    try:
        with DDGS() as ddgs:
            resultados = list(ddgs.text(consulta, max_results=max_results))
        for r in resultados:
            titulo = (r.get("title") or "").strip()
            body = (r.get("body") or "").strip()
            href = (r.get("href") or "").strip()
            if body:
                entry = f"[{titulo}] {body}"
                if href:
                    entry += f" (Fuente: {href})"
                resultados_totales.append(entry)
    except Exception:
        pass

    try:
        resumen_wiki = wikipedia.summary(consulta, sentences=3)
        if resumen_wiki:
            resultados_totales.insert(
                0, f"[Wikipedia] {resumen_wiki}"
            )
    except Exception:
        pass

    if not resultados_totales:
        return "No encontré información sobre ese tema."
    return "\n".join(resultados_totales[:6])


def mensaje_error_amigable(error_texto, contexto=""):
    e = (error_texto or "").lower()
    if "openweather_app_id" in e or "news_api_key" in e or "tmdb_api_key" in e:
        return (
            f"No tengo la API configurada para {contexto or 'esa función'}. "
            f"Puedo intentar una búsqueda web si quieres."
        )
    if "timed out" in e or "connection" in e or "request" in e:
        return "Ahora mismo hay un problema de conexión. Inténtalo en unos segundos."
    return f"No pude completar {contexto or 'esa acción'}."


def dividir_pantalla_izq_der():
    try:
        pyautogui.hotkey("win", "left")
        time.sleep(0.25)
        pyautogui.hotkey("win", "right")
        return "He preparado una distribución de pantalla dividida."
    except Exception as e:
        return f"No pude dividir la pantalla: {e}"


def buscar_archivo_reciente(nombre_parcial="", extension=""):
    try:
        base = Path.home() / "Downloads"
        if not base.exists():
            base = Path.home()
        candidatos = []
        for p in base.rglob("*"):
            if not p.is_file():
                continue
            if extension and p.suffix.lower() != extension.lower():
                continue
            if nombre_parcial and nombre_parcial.lower() not in p.name.lower():
                continue
            candidatos.append(p)
        if not candidatos:
            return "No encontré archivos con ese criterio."
        candidatos.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return f"Encontré: {str(candidatos[0])}"
    except Exception as e:
        return f"No pude buscar archivos: {e}"


def hora_natural_es(dt):
    horas = [
        "doce",
        "una",
        "dos",
        "tres",
        "cuatro",
        "cinco",
        "seis",
        "siete",
        "ocho",
        "nueve",
        "diez",
        "once",
        "doce",
    ]
    minutos_palabra = {
        0: "en punto",
        5: "y cinco",
        10: "y diez",
        15: "y cuarto",
        20: "y veinte",
        25: "y veinticinco",
        30: "y media",
        35: "menos veinticinco",
        40: "menos veinte",
        45: "menos cuarto",
        50: "menos diez",
        55: "menos cinco",
    }
    h24 = dt.hour
    m = dt.minute
    bloque = int(round(m / 5.0) * 5) % 60

    if bloque <= 30:
        h12 = h24 % 12
        frase_hora = horas[h12]
        sufijo = minutos_palabra.get(bloque, f"y {m}")
    else:
        h12 = (h24 + 1) % 12
        frase_hora = horas[h12]
        sufijo = minutos_palabra.get(bloque, f"menos {60 - m}")

    articulo = "Es la" if frase_hora == "una" else "Son las"
    if sufijo == "en punto":
        return f"{articulo} {frase_hora} en punto."
    return f"{articulo} {frase_hora} {sufijo}."


def limpiar_texto(texto):
    return texto.strip().strip('"').strip("'").strip()


def normalizar_ruta(ruta):
    ruta = limpiar_texto(ruta)
    ruta = os.path.expandvars(os.path.expanduser(ruta))
    if not os.path.isabs(ruta):
        ruta = os.path.abspath(ruta)
    return ruta


def normalizar_nombre_app(nombre):
    nombre = limpiar_texto(nombre).lower()
    reemplazos = {
        "el ": "",
        "la ": "",
        "app ": "",
        "aplicacion ": "",
        "aplicación ": "",
        "programa ": "",
        "un ": "",
        "una ": "",
        "mi ": "",
    }
    for prefijo, reemplazo in reemplazos.items():
        if nombre.startswith(prefijo):
            nombre = nombre.replace(prefijo, reemplazo, 1)
    nombre = nombre.strip()
    if nombre.endswith(".exe"):
        nombre_sin_ext = nombre[:-4]
        if nombre_sin_ext in APLICACIONES_CONOCIDAS:
            return nombre_sin_ext
    return nombre


def info_app(nombre):
    nombre_normalizado = normalizar_nombre_app(nombre)
    return APLICACIONES_CONOCIDAS.get(nombre_normalizado, {})


def obtener_comando_lanzamiento(nombre):
    datos = info_app(nombre)
    return datos.get("launch") or limpiar_texto(nombre)


def obtener_procesos_app(nombre):
    datos = info_app(nombre)
    procesos = list(datos.get("processes", []))
    nombre_limpio = limpiar_texto(nombre)
    if nombre_limpio:
        if IS_WINDOWS:
            base = (
                nombre_limpio
                if nombre_limpio.lower().endswith(".exe")
                else f"{nombre_limpio}.exe"
            )
        else:
            base = nombre_limpio.lower()
        if base not in procesos:
            procesos.append(base)
    return procesos


def pulsar_tecla_virtual(vk_code):
    if IS_WINDOWS:
        user32 = windll.user32
        user32.keybd_event(vk_code, 0, 0, 0)
        time.sleep(0.05)
        user32.keybd_event(vk_code, 0, KEYEVENTF_KEYUP, 0)
    else:
        key_map = {
            VK_MEDIA_PLAY_PAUSE: "XF86AudioPlay",
            VK_MEDIA_NEXT_TRACK: "XF86AudioNext",
            VK_MEDIA_PREV_TRACK: "XF86AudioPrev",
            VK_MEDIA_STOP: "XF86AudioStop",
            VK_VOLUME_MUTE: "XF86AudioMute",
            VK_VOLUME_UP: "XF86AudioRaiseVolume",
            VK_VOLUME_DOWN: "XF86AudioLowerVolume",
        }
        xdotool_key = key_map.get(vk_code)
        if xdotool_key:
            try:
                subprocess.run(
                    ["xdotool", "key", xdotool_key],
                    capture_output=True,
                    timeout=5,
                )
            except Exception:
                pass


def extraer_numero(texto):
    numeros = []
    actual = []
    for caracter in texto:
        if caracter.isdigit():
            actual.append(caracter)
        elif actual:
            numeros.append(int("".join(actual)))
            actual = []
    if actual:
        numeros.append(int("".join(actual)))
    return numeros[0] if numeros else None


def mover_ventana_a_monitor(nombre, monitor_idx=1):
    """Move a window to the specified monitor (0=primary, 1=secondary, etc.)."""
    if IS_WINDOWS:
        try:
            import ctypes
            user32 = ctypes.windll.user32

            monitors = []
            def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
                rect = lprcMonitor.contents
                monitors.append((rect.left, rect.top, rect.right, rect.bottom))
                return 1

            MONITORENUMPROC = ctypes.WINFUNCTYPE(
                ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p,
                ctypes.POINTER(ctypes.wintypes.RECT), ctypes.wintypes.LPARAM
            )
            user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0)

            if monitor_idx >= len(monitors):
                return f"Solo tengo {len(monitors)} monitor(es). No puedo mover a monitor {monitor_idx + 1}."

            target = monitors[monitor_idx]
            ventanas = pyautogui.getWindowsWithTitle(limpiar_texto(nombre))
            for v in ventanas:
                if v.title:
                    v.moveTo(target[0], target[1])
                    v.resizeTo(target[2] - target[0], target[3] - target[1])
                    return f"He movido {nombre} al monitor {monitor_idx + 1}."
            return f"No encontré ventana de {nombre} para mover."
        except Exception as e:
            return f"No pude mover {nombre}: {e}"
    else:
        try:
            result = subprocess.run(
                ["xrandr", "--listmonitors"],
                capture_output=True, text=True, timeout=5,
            )
            lines = [l for l in result.stdout.strip().split("\n") if "/" in l]
            if monitor_idx >= len(lines):
                return f"Solo tengo {len(lines)} monitor(es)."
            match = re.search(r"(\d+)/\d+x(\d+)/\d+\+(\d+)\+(\d+)", lines[monitor_idx])
            if match:
                w, h, x, y = match.groups()
                subprocess.run(
                    ["wmctrl", "-r", nombre, "-e", f"0,{x},{y},{w},{h}"],
                    capture_output=True, timeout=5,
                )
                return f"He movido {nombre} al monitor {monitor_idx + 1}."
            return "No pude determinar la posición del monitor."
        except Exception as e:
            return f"No pude mover {nombre}: {e}"


def separar_multi_comandos(texto):
    """Split a sentence with multiple commands into individual commands."""
    texto = texto.strip()
    if not texto:
        return []

    # Patterns that indicate multi-command: "abre X y Y", "abre X, Y y Z"
    # But be careful not to split comparisons like "compara X y Y"
    skip_patterns = [
        r"compar[ae]", r"diferencia", r"mejor.*o\s",
        r"vs\.?", r"versus", r"entre.*y",
    ]
    texto_lower = texto.lower()
    for pattern in skip_patterns:
        if re.search(pattern, texto_lower):
            return [texto]

    # Split on ", y " or " y " when followed by action verbs
    action_verbs = (
        "abre", "abrir", "cierra", "cerrar", "pon", "ponme", "ponlo",
        "mueve", "mover", "busca", "buscar", "reproduce", "ejecuta",
        "activa", "desactiva", "sube", "baja", "silencia", "abre",
    )

    # Try splitting on ", " and " y " / " e "
    separators = [r"\s*,\s*y\s+", r"\s*,\s+", r"\s+y\s+", r"\s+e\s+"]
    for sep in separators:
        parts = re.split(sep, texto, flags=re.IGNORECASE)
        if len(parts) > 1:
            # Check if subsequent parts start with action verbs or could inherit the verb
            comandos = []
            verbo_actual = ""
            for i, part in enumerate(parts):
                part = part.strip()
                if not part:
                    continue
                part_lower = part.lower()
                tiene_verbo = any(part_lower.startswith(v) for v in action_verbs)
                if tiene_verbo or i == 0:
                    comandos.append(part)
                    # Extract the verb from this part
                    for v in action_verbs:
                        if part_lower.startswith(v):
                            verbo_actual = v
                            break
                else:
                    # Check if this part talks about moving to monitor
                    if any(x in part_lower for x in [
                        "segunda pantalla", "segundo monitor", "pantalla secundaria",
                        "monitor secundario", "otra pantalla", "otro monitor",
                    ]):
                        # This is a modifier for the previous command
                        if comandos:
                            comandos.append(f"mueve {_extract_app_from_command(comandos[-1])} a la segunda pantalla")
                        continue
                    # Inherit verb from previous command
                    if verbo_actual:
                        comandos.append(f"{verbo_actual} {part}")
                    else:
                        comandos.append(part)
            if len(comandos) > 1:
                return comandos

    # Check for "ponmelo/ponlo en la segunda pantalla" as a separate action
    monitor_patterns = [
        r"(.+?)\s+(?:y\s+)?(?:ponlo|ponmelo|ponme|muévelo|muevelo|pásalo|pasalo).*?(?:segunda pantalla|segundo monitor|pantalla secundaria|monitor secundario)",
    ]
    for pattern in monitor_patterns:
        match = re.search(pattern, texto, re.IGNORECASE)
        if match:
            base_cmd = match.group(1).strip()
            # Find what app to move
            app_name = _extract_app_from_command(texto)
            if app_name:
                return [base_cmd, f"mueve {app_name} a la segunda pantalla"]

    return [texto]


def _extract_app_from_command(comando):
    """Extract the application name from a command like 'abre spotify'."""
    comando_lower = comando.lower().strip()
    for prefix in ("abre ", "abrir ", "cierra ", "cerrar ", "ejecuta ", "mueve "):
        if comando_lower.startswith(prefix):
            return comando[len(prefix):].strip()
    return comando.strip()


def cerrar_aplicacion(nombre):
    if IS_WINDOWS:
        try:
            ventanas = pyautogui.getWindowsWithTitle(limpiar_texto(nombre))
            for ventana in ventanas:
                if ventana.title:
                    try:
                        ventana.close()
                        return f"He cerrado {nombre}."
                    except Exception:
                        continue
        except Exception:
            pass

        procesos = obtener_procesos_app(nombre)
        for proceso in procesos:
            try:
                resultado = subprocess.run(
                    ["taskkill", "/IM", proceso, "/F"],
                    capture_output=True,
                    text=True,
                    timeout=15,
                )
                if resultado.returncode == 0:
                    return f"He cerrado {nombre}."
            except Exception:
                continue
    else:
        procesos = obtener_procesos_app(nombre)
        for proceso in procesos:
            try:
                resultado = subprocess.run(
                    ["pkill", "-f", proceso],
                    capture_output=True,
                    text=True,
                    timeout=15,
                )
                if resultado.returncode == 0:
                    return f"He cerrado {nombre}."
            except Exception:
                continue

    return f"No encontré un proceso activo para {nombre}."


def activar_ventana(nombre):
    if IS_WINDOWS:
        titulos = info_app(nombre).get("window_titles", [])
        candidatos = titulos + [limpiar_texto(nombre)]
        try:
            for candidato in candidatos:
                if not candidato:
                    continue
                ventanas = pyautogui.getWindowsWithTitle(candidato)
                for ventana in ventanas:
                    if (
                        ventana.title
                        and candidato.lower() in ventana.title.lower()
                    ):
                        try:
                            ventana.activate()
                            return f"He puesto {nombre} al frente."
                        except Exception:
                            continue
        except Exception as e:
            return f"No pude enfocar {nombre}: {e}"
    else:
        try:
            subprocess.run(
                ["wmctrl", "-a", nombre],
                capture_output=True,
                timeout=5,
            )
            return f"He puesto {nombre} al frente."
        except Exception:
            pass
    return f"No encontré una ventana abierta de {nombre}."


def abrir_aplicacion(nombre):
    nombre_clean = nombre.strip().strip('"').strip("'")
    cmd = obtener_comando_lanzamiento(nombre_clean)

    if IS_WINDOWS:
        # Try os.startfile first (handles .exe, .lnk, URLs, etc.)
        try:
            os.startfile(cmd)
            return f"He abierto {nombre_clean}."
        except Exception:
            pass
        # Try finding the .exe in common paths
        if not cmd.endswith(".exe"):
            exe_name = cmd + ".exe"
        else:
            exe_name = cmd
        search_paths = [
            os.path.join(os.environ.get("PROGRAMFILES", "C:\\Program Files"), "**", exe_name),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"), "**", exe_name),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "**", exe_name),
            os.path.join(os.environ.get("APPDATA", ""), "**", exe_name),
        ]
        for pattern in search_paths:
            if not pattern:
                continue
            matches = glob.glob(pattern, recursive=True)
            if matches:
                try:
                    os.startfile(matches[0])
                    return f"He abierto {nombre_clean}."
                except Exception:
                    continue
        # Try Start-Process as last resort
        try:
            subprocess.Popen(
                ["powershell", "-NoProfile", "-Command", f"Start-Process '{cmd}'"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            return f"He lanzado {nombre_clean}."
        except Exception:
            pass

    # Linux / generic fallback
    try:
        subprocess.Popen(
            [cmd],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return f"He abierto {nombre_clean}."
    except Exception:
        try:
            subprocess.Popen(
                cmd,
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return f"He lanzado {nombre_clean}."
        except Exception as e:
            return f"No pude abrir {nombre_clean}: {e}"


def _obtener_dispositivo_audio():
    if not HAS_WIN_AUDIO:
        raise OSError("Control de audio nativo solo disponible en Windows.")
    ole32 = windll.ole32
    ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
    enumerator = POINTER(IMMDeviceEnumerator)()
    hr = ole32.CoCreateInstance(
        byref(GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")),
        None,
        CLSCTX_ALL,
        byref(IMMDeviceEnumerator._iid_),
        byref(enumerator),
    )
    if hr != 0:
        raise OSError(f"CoCreateInstance fallo con codigo {hr}")
    vtbl = enumerator.contents.lpVtbl.contents
    endpoint = POINTER(IMMDevice)()
    hr = vtbl.GetDefaultAudioEndpoint(enumerator, 0, 1, byref(endpoint))
    if hr != 0:
        raise OSError(f"GetDefaultAudioEndpoint fallo con codigo {hr}")
    return endpoint


def _obtener_endpoint_volume():
    endpoint = _obtener_dispositivo_audio()
    volume = POINTER(IAudioEndpointVolume)()
    hr = endpoint.contents.lpVtbl.contents.Activate(
        endpoint,
        byref(IAudioEndpointVolume._iid_),
        CLSCTX_ALL,
        None,
        byref(volume),
    )
    if hr != 0:
        raise OSError(f"Activate fallo con codigo {hr}")
    return volume


def obtener_volumen_actual():
    if not HAS_WIN_AUDIO:
        try:
            result = subprocess.run(
                ["amixer", "get", "Master"],
                capture_output=True,
                text=True,
                timeout=5,
            )
            match = re.search(r"\[(\d+)%\]", result.stdout)
            if match:
                return int(match.group(1))
        except Exception:
            pass
        return 50
    volume = _obtener_endpoint_volume()
    nivel = c_float()
    hr = volume.contents.lpVtbl.contents.GetMasterVolumeLevelScalar(
        volume, byref(nivel)
    )
    if hr != 0:
        raise OSError(f"GetMasterVolumeLevelScalar fallo con codigo {hr}")
    return round(nivel.value * 100)


def fijar_volumen(porcentaje):
    porcentaje = max(0, min(100, int(porcentaje)))
    if not HAS_WIN_AUDIO:
        try:
            subprocess.run(
                ["amixer", "set", "Master", f"{porcentaje}%"],
                capture_output=True,
                timeout=5,
            )
            return f"Volumen ajustado al {porcentaje}%."
        except Exception:
            return f"No pude ajustar el volumen en este sistema."
    volume = _obtener_endpoint_volume()
    hr = volume.contents.lpVtbl.contents.SetMasterVolumeLevelScalar(
        volume, c_float(porcentaje / 100), None
    )
    if hr != 0:
        raise OSError(f"SetMasterVolumeLevelScalar fallo con codigo {hr}")
    return f"Volumen ajustado al {porcentaje}%."


def cambiar_volumen(delta):
    actual = obtener_volumen_actual()
    destino = max(0, min(100, actual + delta))
    return fijar_volumen(destino)


def silenciar_audio(silencio=True):
    if not HAS_WIN_AUDIO:
        try:
            cmd = "mute" if silencio else "unmute"
            subprocess.run(
                ["amixer", "set", "Master", cmd],
                capture_output=True,
                timeout=5,
            )
            return "Audio silenciado." if silencio else "Audio activado."
        except Exception:
            return "No pude controlar el audio en este sistema."
    volume = _obtener_endpoint_volume()
    hr = volume.contents.lpVtbl.contents.SetMute(
        volume, bool(silencio), None
    )
    if hr != 0:
        raise OSError(f"SetMute fallo con codigo {hr}")
    return "Audio silenciado." if silencio else "Audio activado."


def controlar_multimedia(accion):
    mapa = {
        "play_pause": (VK_MEDIA_PLAY_PAUSE, "Play/Pausa ejecutado."),
        "next": (VK_MEDIA_NEXT_TRACK, "Siguiente pista."),
        "previous": (VK_MEDIA_PREV_TRACK, "Pista anterior."),
        "stop": (VK_MEDIA_STOP, "Audio detenido."),
    }
    vk_code, mensaje = mapa[accion]
    pulsar_tecla_virtual(vk_code)
    return mensaje


def controlar_video(accion):
    if IS_WINDOWS:
        browser_titles = ["google chrome", "chrome", "opera", "opera gx"]
        for title in browser_titles:
            try:
                windows = pyautogui.getWindowsWithTitle(title)
                if windows:
                    windows[0].activate()
                    time.sleep(0.1)
                    break
            except Exception:
                continue

    if accion == "fullscreen":
        pyautogui.press("f")
        return "Video puesto en pantalla completa."
    elif accion == "exit_fullscreen":
        pyautogui.press("f")
        return "Salido de pantalla completa."
    elif accion == "play":
        pyautogui.press("space")
        return "Video iniciado."
    elif accion == "pause":
        pyautogui.press("space")
        return "Video pausado."
    return "Acción no reconocida."


def mover_raton(x, y):
    pyautogui.moveTo(x, y)
    return f"Ratón movido a {x}, {y}."


def click_raton(boton="left"):
    pyautogui.click(button=boton)
    return f"Click {boton} ejecutado."


def doble_click():
    pyautogui.doubleClick()
    return "Doble click ejecutado."


def scroll_raton(direccion, cantidad=3):
    if "arriba" in direccion.lower() or "up" in direccion.lower():
        pyautogui.scroll(cantidad)
    else:
        pyautogui.scroll(-cantidad)
    return f"Scroll {direccion} ejecutado."


def abrir_recurso(destino):
    destino = limpiar_texto(destino)
    if not destino:
        return "No me dijiste qué abrir."

    if destino.startswith(("http://", "https://")):
        webbrowser.open(destino)
        return f"Abriendo URL: {destino}"

    ruta = normalizar_ruta(destino)
    if os.path.exists(ruta):
        if IS_WINDOWS:
            os.startfile(ruta)
        else:
            subprocess.Popen(
                ["xdg-open", ruta],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        return f"Abierto: {ruta}"

    try:
        subprocess.Popen(
            destino,
            shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return f"He lanzado: {destino}"
    except Exception as e:
        return f"No pude abrir '{destino}': {e}"


def ejecutar_comando_sistema(comando, como_admin=False):
    """Cross-platform command execution with optional admin elevation."""
    try:
        if IS_WINDOWS:
            if como_admin:
                cmd_escaped = comando.replace("'", "''")
                resultado = subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     f"Start-Process powershell -ArgumentList '-NoProfile','-Command','{cmd_escaped}' -Verb RunAs -Wait"],
                    capture_output=True, text=True, timeout=60,
                )
            else:
                resultado = subprocess.run(
                    ["powershell", "-NoProfile", "-Command", comando],
                    capture_output=True, text=True, timeout=30,
                )
        else:
            resultado = subprocess.run(
                ["bash", "-c", comando],
                capture_output=True, text=True, timeout=30,
            )
    except subprocess.TimeoutExpired:
        return "El comando tardó demasiado y fue cancelado."
    except Exception as e:
        return f"No pude ejecutar el comando: {e}"

    salida = (resultado.stdout or "").strip()
    error = (resultado.stderr or "").strip()

    if resultado.returncode == 0:
        return salida or "Comando ejecutado correctamente."

    if "acceso denegado" in error.lower() or "access denied" in error.lower():
        if como_admin:
            return f"Acceso denegado incluso con permisos de administrador: {error}"
        return f"Acceso denegado. Para ejecutar este comando con permisos de administrador, diga: ejecuta como admin {comando}"

    if error:
        return f"El comando falló: {error}"
    return f"El comando falló con código {resultado.returncode}."


def crear_archivo(ruta, contenido=""):
    ruta = normalizar_ruta(ruta)
    carpeta = os.path.dirname(ruta)
    if carpeta:
        os.makedirs(carpeta, exist_ok=True)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)
    return f"Archivo creado: {ruta}"


def sobrescribir_archivo(ruta, contenido):
    ruta = normalizar_ruta(ruta)
    carpeta = os.path.dirname(ruta)
    if carpeta:
        os.makedirs(carpeta, exist_ok=True)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)
    return f"Archivo actualizado: {ruta}"


def agregar_a_archivo(ruta, contenido):
    ruta = normalizar_ruta(ruta)
    carpeta = os.path.dirname(ruta)
    if carpeta:
        os.makedirs(carpeta, exist_ok=True)
    with open(ruta, "a", encoding="utf-8") as f:
        f.write(contenido)
    return f"Contenido agregado a: {ruta}"


def leer_archivo(ruta, limite=4000):
    ruta = normalizar_ruta(ruta)
    if not os.path.exists(ruta):
        return f"No existe el archivo: {ruta}"
    with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
        contenido = f.read(limite)
    return f"Contenido de {ruta}:\n{contenido}"


def leer_documento(ruta, limite=12000):
    ruta = normalizar_ruta(ruta)
    if not os.path.exists(ruta):
        return None, f"No existe el archivo: {ruta}"

    extension = Path(ruta).suffix.lower()
    try:
        if extension == ".pdf":
            if PdfReader is None:
                return (
                    None,
                    "No puedo leer PDF porque falta la librería 'pypdf'. Instala: pip install pypdf",
                )
            reader = PdfReader(ruta)
            texto = []
            for pagina in reader.pages:
                texto.append(pagina.extract_text() or "")
            contenido = "\n".join(texto).strip()
        else:
            with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
                contenido = f.read()
    except Exception as e:
        return None, f"No pude leer el archivo: {e}"

    if not contenido.strip():
        return None, f"El archivo {ruta} no contiene texto legible."

    return contenido[:limite], None


def resumir_documento(ruta):
    contenido, error = leer_documento(ruta)
    if error:
        return error
    prompt = (
        "Resume este documento en español con: 1) objetivo, 2) pasos de ejecución, "
        "3) checklist accionable.\n\nDOCUMENTO:\n"
        f"{contenido}"
    )
    return consultar_ia(prompt, cargar_memoria())


def generar_plan_desde_documento(ruta):
    contenido, error = leer_documento(ruta)
    if error:
        return error
    prompt = (
        "Analiza estas instrucciones y conviértelas en acciones ejecutables. "
        "Devuelve cada paso en formato ACTION: <accion>. Lista las acciones en orden.\n\n"
        f"INSTRUCCIONES:\n{contenido}"
    )
    plan = consultar_ia(prompt, cargar_memoria())
    ejecutar_plan(plan)
    return f"Plan generado y ejecutado:\n{plan}"


def responder_preguntas_de_documento(ruta):
    contenido, error = leer_documento(ruta, limite=30000)
    if error:
        return error
    prompt = (
        "Lee el siguiente documento en español. Extrae TODAS las preguntas que aparezcan y respóndelas "
        "de forma clara y breve. Si alguna pregunta no tiene contexto suficiente en el documento, indícalo.\n\n"
        f"DOCUMENTO:\n{contenido}"
    )
    return consultar_ia(prompt, cargar_memoria())


def cambiar_directorio(path):
    try:
        os.chdir(path)
        return f"Directorio cambiado a {os.getcwd()}"
    except Exception as e:
        return f"Error cambiando directorio: {e}"


def listar_archivos(path="."):
    try:
        files = os.listdir(path)
        return "\n".join(files)
    except Exception as e:
        return f"Error listando archivos: {e}"


def crear_directorio(name):
    try:
        os.makedirs(name, exist_ok=True)
        return f"Directorio {name} creado."
    except Exception as e:
        return f"Error creando directorio: {e}"


def borrar_directorio(name):
    try:
        os.rmdir(name)
        return f"Directorio {name} borrado."
    except Exception as e:
        return f"Error borrando directorio: {e}"


def mover_archivo(origen, destino):
    try:
        shutil.move(origen, destino)
        return f"Archivo movido de {origen} a {destino}."
    except Exception as e:
        return f"Error moviendo archivo: {e}"


def copiar_archivo(origen, destino):
    try:
        shutil.copy(origen, destino)
        return f"Archivo copiado de {origen} a {destino}."
    except Exception as e:
        return f"Error copiando archivo: {e}"


def escribir_en_archivo(ruta, contenido, modo="w"):
    try:
        with open(ruta, modo, encoding="utf-8") as f:
            f.write(contenido)
        return f"Contenido escrito en {ruta}."
    except Exception as e:
        return f"Error escribiendo en archivo: {e}"


ACCIONES_AUTOMATIZADAS = {
    "open_intellij": lambda: abrir_aplicacion("intellij"),
    "open_vs_code": lambda: abrir_aplicacion("vs code"),
    "write_code": lambda code: escribir_codigo_en_editor(code),
    "run_command": lambda cmd: ejecutar_comando_sistema(cmd),
    "open_mysql_workbench": lambda: abrir_aplicacion("mysql workbench"),
    "open_dbeaver": lambda: abrir_aplicacion("dbeaver"),
    "open_windsurf": lambda: abrir_aplicacion("windsurf"),
}


def leer_pantalla():
    if pytesseract is None:
        return "OCR no disponible. Instala pytesseract y Tesseract-OCR."
    screenshot = pyautogui.screenshot()
    text = pytesseract.image_to_string(screenshot, lang="spa")
    return text.strip()


def activar_inicio_automatico():
    if not IS_WINDOWS:
        return "Inicio automático solo disponible en Windows por ahora."
    try:
        import winreg

        ruta_script = os.path.abspath("main.py")
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        )
        winreg.SetValueEx(key, "Jarvis", 0, winreg.REG_SZ, f'py "{ruta_script}"')
        winreg.CloseKey(key)
        return "Inicio automático activado. Jarvis se ejecutará al encender el PC."
    except Exception as e:
        return f"Error activando inicio automático: {e}"


def desactivar_inicio_automatico():
    if not IS_WINDOWS:
        return "Inicio automático solo disponible en Windows por ahora."
    try:
        import winreg

        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        )
        winreg.DeleteValue(key, "Jarvis")
        winreg.CloseKey(key)
        return "Inicio automático desactivado."
    except Exception as e:
        return f"Error desactivando inicio automático: {e}"


def ejecutar_plan(plan):
    lines = plan.split("\n")
    for line in lines:
        line = line.strip()
        if line.startswith("ACTION: "):
            action_str = line[len("ACTION: ") :].strip()
            if " " in action_str:
                parts = action_str.split(" ", 1)
                cmd = parts[0]
                arg = parts[1]
                if cmd in ACCIONES_AUTOMATIZADAS:
                    try:
                        ACCIONES_AUTOMATIZADAS[cmd](arg)
                        time.sleep(1)
                    except Exception:
                        pass
            else:
                if action_str in ACCIONES_AUTOMATIZADAS:
                    try:
                        ACCIONES_AUTOMATIZADAS[action_str]()
                        time.sleep(1)
                    except Exception:
                        pass


def capturar_pantalla():
    os.makedirs(CARPETA_CAPTURAS, exist_ok=True)
    nombre = f"captura_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    ruta = os.path.join(CARPETA_CAPTURAS, nombre)
    imagen = pyautogui.screenshot()
    imagen.save(ruta)
    return f"Captura guardada en: {ruta}"


def obtener_ventana_activa():
    try:
        ventana = pyautogui.getActiveWindow()
        if not ventana:
            return "No detecto una ventana activa."
        titulo = ventana.title or "sin titulo"
        return f"Ventana activa: {titulo}"
    except Exception as e:
        return f"No pude leer la ventana activa: {e}"


def escribir_en_ventana(texto):
    texto = texto.strip()
    if not texto:
        return "No me diste texto para escribir."
    pyautogui.write(texto, interval=0.01)
    return "He escrito en la ventana activa."


def pulsar_teclas(secuencia):
    partes = [
        limpiar_texto(x) for x in secuencia.replace("+", " ").split() if x.strip()
    ]
    if not partes:
        return "No me diste teclas para pulsar."
    pyautogui.hotkey(*partes)
    return f"Atajo ejecutado: {' + '.join(partes)}"


def presionar_tecla(tecla):
    tecla = limpiar_texto(tecla).lower()
    if not tecla:
        return "No me dijiste qué tecla pulsar."
    pyautogui.press(tecla)
    return f"Tecla pulsada: {tecla}"


def escribir_codigo_en_editor(codigo):
    codigo = codigo.strip()
    if not codigo:
        return "No me diste código para escribir."
    pyautogui.write(codigo, interval=0.005)
    return "Código escrito en la ventana activa."


def copiar_texto_al_portapapeles(texto):
    try:
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(texto)
        root.update()
        root.destroy()
        return True
    except Exception:
        return False


def pegar_texto_en_ventana(texto):
    texto = texto.strip()
    if not texto:
        return "No me diste texto para pegar."
    if copiar_texto_al_portapapeles(texto):
        pyautogui.hotkey("ctrl", "v")
        return "He pegado el texto en la ventana activa."
    pyautogui.write(texto, interval=0.01)
    return "He escrito el texto en la ventana activa."


def seleccionar_todo():
    pyautogui.hotkey("ctrl", "a")
    return "He seleccionado todo."


def reemplazar_seleccion(texto):
    texto = texto.strip()
    if not texto:
        return "No me diste texto para reemplazar."
    return pegar_texto_en_ventana(texto)


def nueva_linea(cantidad=1):
    cantidad = max(1, int(cantidad))
    for _ in range(cantidad):
        pyautogui.press("enter")
    return "He añadido una nueva línea."


def tabular(cantidad=1):
    cantidad = max(1, int(cantidad))
    for _ in range(cantidad):
        pyautogui.press("tab")
    return "He insertado tabulación."


def borrar_seleccion():
    pyautogui.press("backspace")
    return "He borrado la selección."


def cortar():
    pyautogui.hotkey("ctrl", "x")
    return "He cortado la selección."


def rehacer():
    pyautogui.hotkey("ctrl", "y")
    return "He rehecho la última acción."


def ir_a_url(destino):
    destino = limpiar_texto(destino)
    if not destino:
        return "No me dijiste a qué página ir."
    if not destino.startswith(("http://", "https://")):
        if "." in destino and " " not in destino:
            destino = f"https://{destino}"
        else:
            destino = f"https://www.google.com/search?q={quote_plus(destino)}"
    webbrowser.open(destino)
    return f"Abriendo {destino}."


def buscar_en_sitio_web(sitio, consulta):
    sitio = limpiar_texto(sitio).lower()
    consulta = limpiar_texto(consulta)
    if not sitio or not consulta:
        return "Necesito el sitio y la búsqueda."

    plantillas = {
        "youtube": "https://www.youtube.com/results?search_query={q}",
        "google": "https://www.google.com/search?q={q}",
        "github": "https://github.com/search?q={q}",
        "wikipedia": "https://es.wikipedia.org/w/index.php?search={q}",
        "amazon": "https://www.amazon.es/s?k={q}",
        "reddit": "https://www.reddit.com/search/?q={q}",
        "x": "https://x.com/search?q={q}",
        "twitter": "https://x.com/search?q={q}",
        "ebay": "https://www.ebay.es/sch/i.html?_nkw={q}",
        "aliexpress": "https://www.aliexpress.com/wholesale?SearchText={q}",
        "stackoverflow": "https://stackoverflow.com/search?q={q}",
    }
    plantilla = plantillas.get(sitio)
    if not plantilla:
        return ir_a_url(f"site:{sitio} {consulta}")
    url = plantilla.format(q=quote_plus(consulta))
    webbrowser.open(url)
    return f"Buscando {consulta} en {sitio}."


def extraer_consulta_busqueda(comando, disparadores):
    consulta = limpiar_texto(comando)
    consulta_lower = consulta.lower()

    for disparador in disparadores:
        disparador_lower = disparador.lower()
        if disparador_lower in consulta_lower:
            indice = consulta_lower.find(disparador_lower)
            consulta = consulta[indice + len(disparador) :].strip()
            consulta_lower = consulta.lower()
            break

    prefijos_ruido = [
        "muéstrame ",
        "muestrame ",
        "ponme ",
        "pon ",
        "busca ",
        "búscame ",
        "buscame ",
        "quiero ver ",
        "quiero escuchar ",
        "abre ",
        "reproduce ",
        "pon el ",
        "pon la ",
        "el ",
        "la ",
        "los ",
        "las ",
        "un ",
        "una ",
    ]
    for prefijo in prefijos_ruido:
        while consulta_lower.startswith(prefijo):
            consulta = consulta[len(prefijo) :].strip()
            consulta_lower = consulta.lower()

    sufijos_ruido = [
        " en youtube",
        " por youtube",
        " en google",
    ]
    for sufijo in sufijos_ruido:
        while consulta_lower.endswith(sufijo):
            consulta = consulta[: -len(sufijo)].strip()
            consulta_lower = consulta.lower()

    while consulta and consulta[0] in ",:;.-":
        consulta = consulta[1:].strip()
    while consulta and consulta[-1] in ",:;.-":
        consulta = consulta[:-1].strip()

    return consulta


def elegir_voz_masculina(engine):
    try:
        voices = engine.getProperty("voices") or []
    except Exception:
        return None

    priorities = [
        ("david", ""),
        ("pablo", ""),
        ("hector", ""),
        ("raul", ""),
        ("mateo", ""),
        ("spanish", "male"),
        ("es_", "male"),
        ("spanish", "mascul"),
        ("es-", "mascul"),
        ("male", ""),
        ("mascul", ""),
    ]

    def normalizar_voz(voz):
        voz_id = getattr(voz, "id", "") or ""
        voz_name = getattr(voz, "name", "") or ""
        partes = [str(voz_id), str(voz_name)]
        idiomas = getattr(voz, "languages", []) or []
        for idioma in idiomas:
            if isinstance(idioma, bytes):
                try:
                    partes.append(idioma.decode("utf-8", errors="ignore"))
                except Exception:
                    continue
            else:
                partes.append(str(idioma))
        return " ".join(partes).lower()

    normalized_voices = [(voz, normalizar_voz(voz)) for voz in voices]
    for language, gender in priorities:
        for voz, texto in normalized_voices:
            if language and language not in texto:
                continue
            if gender and gender not in texto:
                continue
            voz_id = getattr(voz, "id", None)
            if voz_id:
                return voz_id

    for voz, texto in normalized_voices:
        if "spanish" in texto or "es_" in texto or "es-" in texto:
            voz_id = getattr(voz, "id", None)
            if voz_id:
                return voz_id
    if voices:
        voz_id = getattr(voices[0], "id", None)
        return voz_id
    return None


def dirigir_como_senor(texto):
    texto = (texto or "").strip()
    if not texto:
        return "Listo."
    texto = re.sub(r"\s+", " ", texto).strip()
    texto = re.sub(
        r"^(lo siento|lo lamento|disculpa|perd[oó]n|perdona|siento decirte)\b[,:\s-]*",
        "",
        texto,
        flags=re.IGNORECASE,
    ).strip()
    texto = re.sub(r"^\s*(señor[,:\s-]*)+", "", texto, flags=re.IGNORECASE).strip()
    texto = re.sub(r"([, ]+señor[.!?]?)$", "", texto, flags=re.IGNORECASE).strip()
    return texto or "Listo."


def preparar_contexto_modelo(historial, limite=12):
    base = []
    if historial and isinstance(historial, list):
        for msg in historial:
            if msg.get("role") == "system":
                base.append(msg)
                break
    if not base:
        base = [
            {
                "role": "system",
                "content": (
                    "Eres Jarvis, un asistente inteligente estilo Iron Man. "
                    "Responde en español claro, natural y breve. "
                    "Puedes hacer comparativas, buscar información y dar datos concretos. "
                    "No uses relleno ni frases repetitivas."
                ),
            }
        ]
    conversacion = [
        m for m in historial if m.get("role") in {"user", "assistant"}
    ][-limite:]
    return base + conversacion


def formatear_lista(items):
    items = [str(item).strip() for item in items if str(item).strip()]
    if not items:
        return "No encontré resultados."
    return "\n".join(f"- {item}" for item in items)


def extraer_ciudad_clima(comando):
    texto = comando.strip()
    texto_lower = texto.lower()
    for trigger in (
        "clima en ",
        "tiempo en ",
        "temperatura en ",
        "weather in ",
        "weather en ",
    ):
        if trigger in texto_lower:
            index = texto_lower.find(trigger)
            return texto[index + len(trigger) :].strip(" .,:;")
    return ""


def extraer_whatsapp(comando):
    texto = limpiar_texto(comando)
    texto_lower = texto.lower()
    triggers = (
        "envia un whatsapp al ",
        "envía un whatsapp al ",
        "manda un whatsapp al ",
        "manda whatsapp al ",
        "send whatsapp message to ",
    )
    for trigger in triggers:
        if trigger in texto_lower:
            start = texto_lower.find(trigger) + len(trigger)
            rest = texto[start:].strip()
            for separator in (
                " diciendo ",
                " con el mensaje ",
                " que diga ",
                " mensaje ",
            ):
                idx = rest.lower().find(separator)
                if idx != -1:
                    return rest[:idx].strip(), rest[idx + len(separator) :].strip()
            return rest, ""
    return "", ""


def extraer_email(comando):
    texto = limpiar_texto(comando)
    texto_lower = texto.lower()
    triggers = (
        "envia un correo a ",
        "envía un correo a ",
        "manda un correo a ",
        "send an email to ",
    )
    for trigger in triggers:
        if trigger in texto_lower:
            start = texto_lower.find(trigger) + len(trigger)
            rest = texto[start:].strip()
            subject_sep = " asunto "
            message_sep = " mensaje "
            subject_idx = rest.lower().find(subject_sep)
            message_idx = rest.lower().find(message_sep)
            if (
                subject_idx != -1
                and message_idx != -1
                and message_idx > subject_idx
            ):
                return (
                    rest[:subject_idx].strip(),
                    rest[subject_idx + len(subject_sep) : message_idx].strip(),
                    rest[message_idx + len(message_sep) :].strip(),
                )
            return rest, "", ""
    return "", "", ""


def ejecutar_accion_local(comando):
    comando_original = comando.strip()
    comando_lower = comando_original.lower()

    if any(
        x in comando_lower
        for x in [
            "que hora es",
            "qué hora es",
            "dime la hora",
            "hora actual",
            "que hora tenemos",
            "qué hora tenemos",
        ]
    ):
        ahora = datetime.now()
        return hora_natural_es(ahora)

    if any(
        x in comando_lower
        for x in [
            "que dia es",
            "qué día es",
            "fecha de hoy",
            "que fecha es",
            "qué fecha es",
        ]
    ):
        hoy = datetime.now()
        dias = [
            "lunes",
            "martes",
            "miércoles",
            "jueves",
            "viernes",
            "sábado",
            "domingo",
        ]
        meses = [
            "enero",
            "febrero",
            "marzo",
            "abril",
            "mayo",
            "junio",
            "julio",
            "agosto",
            "septiembre",
            "octubre",
            "noviembre",
            "diciembre",
        ]
        return f"Hoy es {dias[hoy.weekday()]}, {hoy.day} de {meses[hoy.month - 1]} de {hoy.year}."

    if any(
        x in comando_lower
        for x in [
            "despiertame a las",
            "despiértame a las",
            "pon una alarma",
            "alarma para",
        ]
    ):
        return crear_alarma(comando_original)

    if "temporizador" in comando_lower:
        return crear_temporizador(comando_original)

    if (
        any(x in comando_lower for x in ["añade", "agrega", "crear evento"])
        and "calendario" in comando_lower
    ):
        return crear_evento_calendario(comando_original)

    if comando_lower.startswith("recuerdame") or comando_lower.startswith(
        "recuérdame"
    ):
        return crear_recordatorio(comando_original)

    # Comparativas
    tema1, tema2 = detectar_comparativa(comando_original)
    if tema1 and tema2:
        datos_web = realizar_comparativa_web(tema1, tema2)
        prompt = (
            f"El usuario quiere comparar '{tema1}' con '{tema2}'. "
            f"Aquí tienes datos de internet:\n{datos_web}\n\n"
            f"Haz una comparativa clara y estructurada con datos concretos. "
            f"Usa formato de tabla si es posible. Responde en español."
        )
        return consultar_ia_directa(prompt)

    if any(
        x in comando_lower
        for x in ["trafico", "tráfico"]
    ) and any(x in comando_lower for x in ["centro", "llegar", "tardare", "tardaré"]):
        return f"Tráfico en tiempo real:\n{obtener_resumen_web('tráfico actual para llegar al centro en coche hoy', 3)}"

    if any(
        x in comando_lower
        for x in ["quedo el", "quedó el", "resultado del", "como quedo", "cómo quedó"]
    ):
        consulta = comando_original.replace("¿", "").replace("?", "").strip()
        return f"Resultado deportivo:\n{obtener_resumen_web(consulta, 3)}"

    if any(
        x in comando_lower
        for x in [
            "noticias mas importantes",
            "noticias más importantes",
            "top noticias",
            "3 noticias",
        ]
    ):
        try:
            noticias = get_latest_news()
            top = noticias[:3]
            return f"Top 3 noticias de hoy:\n{formatear_lista(top)}"
        except Exception:
            return f"Top 3 noticias (web):\n{obtener_resumen_web('3 noticias más importantes de hoy', 3)}"

    if any(
        x in comando_lower
        for x in [
            "lee mis notificaciones",
            "léeme los mensajes nuevos",
            "notificaciones nuevas",
        ]
    ):
        return "Aún no tengo acceso nativo a notificaciones del sistema. Puedo abrir WhatsApp Web o Gmail para revisarlas contigo."

    if any(
        x in comando_lower for x in ["busca el correo", "buscar correo"]
    ) and any(x in comando_lower for x in ["codigo", "código"]):
        query = quote_plus(
            comando_original.replace("busca", "").replace("buscar", "").strip()
        )
        webbrowser.open(f"https://mail.google.com/mail/u/0/#search/{query}")
        return "He abierto Gmail con una búsqueda del correo solicitado."

    if any(
        x in comando_lower
        for x in [
            "cuantos euros",
            "cuántos euros",
            "grados",
            "fahrenheit",
            "celsius",
            "dolares",
            "dólares",
        ]
    ):
        conversion = convertir_unidades(comando_original)
        if conversion:
            return conversion

    if any(x in comando_lower for x in ["como se dice", "cómo se dice", "traduce"]):
        return traducir_texto(comando_original)

    if any(x in comando_lower for x in ["calcula", "porcentaje", "% de"]):
        calculo = calcular_rapido(comando_original)
        if calculo:
            return calculo

    if "podcast" in comando_lower:
        webbrowser.open("https://open.spotify.com/genre/podcasts-web")
        return "He abierto Spotify en la sección de podcasts."

    if any(x in comando_lower for x in ["dato curioso", "curiosidad"]):
        consulta = comando_original.replace("dime", "").replace("sobre", "").strip()
        return f"Dato curioso:\n{obtener_resumen_web(f'dato curioso {consulta}', 1)}"

    if any(x in comando_lower for x in ["playlist", "spotify"]):
        nombre = ""
        m = re.search(r"[\"']([^\"']+)[\"']", comando_original)
        if m:
            nombre = m.group(1).strip()
        elif "playlist" in comando_lower:
            partes = comando_original.lower().split("playlist", 1)[1].strip(" de:;,.")
            nombre = partes.strip()
        if nombre:
            webbrowser.open(f"https://open.spotify.com/search/{quote_plus(nombre)}")
            return f"He buscado '{nombre}' en Spotify."

    if any(
        x in comando_lower
        for x in [
            "pon opera a la izquierda",
            "split screen",
            "pantalla dividida",
        ]
    ):
        return dividir_pantalla_izq_der()

    if (
        any(
            x in comando_lower
            for x in ["donde esta el pdf", "dónde está el pdf", "factura"]
        )
        and "ayer" in comando_lower
    ):
        return buscar_archivo_reciente("factura", ".pdf")

    if (
        any(
            x in comando_lower
            for x in ["busca el codigo de error", "busca el código de error"]
        )
        and "log" in comando_lower
    ):
        return buscar_archivo_reciente("log", "")

    if any(x in comando_lower for x in ["cpu", "temperatura", "calentando"]):
        if IS_WINDOWS:
            return ejecutar_comando_sistema(
                "(Get-Counter '\\Processor(_Total)\\% Processor Time').CounterSamples.CookedValue"
            )
        else:
            return ejecutar_comando_sistema("top -bn1 | head -5")

    try:
        if any(
            x in comando_lower
            for x in [
                "mi direccion ip",
                "mi dirección ip",
                "cual es mi ip",
                "cuál es mi ip",
                "ip address",
            ]
        ):
            try:
                return f"Tu dirección IP pública es {find_my_ip()}."
            except Exception as e:
                return mensaje_error_amigable(str(e), "consultar tu IP")

        if "wikipedia" in comando_lower:
            termino = extraer_consulta_busqueda(
                comando_original,
                [
                    "busca en wikipedia",
                    "buscar en wikipedia",
                    "wikipedia",
                    "en wikipedia",
                ],
            )
            if not termino:
                return "No entendí qué quieres buscar en Wikipedia."
            try:
                return f"Según Wikipedia: {search_on_wikipedia(termino)}"
            except Exception as e:
                return mensaje_error_amigable(
                    str(e), f"buscar en Wikipedia '{termino}'"
                )

        if any(
            x in comando_lower
            for x in [
                "noticias",
                "últimas noticias",
                "ultimas noticias",
                "news",
                "titulares",
            ]
        ):
            try:
                noticias = get_latest_news()
                return f"Estos son los últimos titulares:\n{formatear_lista(noticias)}"
            except Exception:
                return f"Noticias (web):\n{obtener_resumen_web('últimas noticias España hoy', 3)}"

        if any(
            x in comando_lower
            for x in [
                "peliculas en tendencia",
                "películas en tendencia",
                "trending movies",
                "peliculas trending",
                "películas trending",
            ]
        ):
            try:
                peliculas = get_trending_movies()
                return f"Estas son algunas películas en tendencia:\n{formatear_lista(peliculas)}"
            except Exception as e:
                fallback = obtener_resumen_web("películas en tendencia hoy", 3)
                if "No pude consultar" not in fallback:
                    return f"Películas en tendencia (web):\n{fallback}"
                return mensaje_error_amigable(
                    str(e), "obtener películas en tendencia"
                )

        if any(x in comando_lower for x in ["chiste", "joke", "humor negro"]):
            try:
                if any(
                    x in comando_lower
                    for x in [
                        "humor negro",
                        "chiste negro",
                        "dark joke",
                    ]
                ):
                    return get_random_joke("negro")
                if any(
                    x in comando_lower for x in ["chiste bueno", "buen chiste"]
                ):
                    return get_random_joke("bueno")
                if any(x in comando_lower for x in ["chiste malo", "mal chiste"]):
                    return get_random_joke("malo")
                return get_random_joke("normal")
            except Exception as e:
                return mensaje_error_amigable(str(e), "obtener un chiste")

        if any(
            x in comando_lower
            for x in [
                "dame un consejo",
                "quiero un consejo",
                "advice",
                "necesito un consejo",
            ]
        ):
            try:
                return get_random_advice()
            except Exception as e:
                return mensaje_error_amigable(str(e), "obtener un consejo")

        if (
            any(x in comando_lower for x in ["clima", "tiempo", "temperatura"])
            and "volumen" not in comando_lower
        ):
            ciudad = extraer_ciudad_clima(comando_original)
            if not ciudad:
                try:
                    ciudad = get_city_from_ip()
                except Exception:
                    ciudad = "Madrid"
            try:
                weather, temperature, feels_like = get_weather_report(ciudad)
                return (
                    f"El clima en {ciudad} es {weather}. "
                    f"La temperatura actual es {temperature} y se siente como {feels_like}."
                )
            except Exception as e:
                fallback = obtener_resumen_web(f"clima actual en {ciudad}", 1)
                if "No pude consultar" not in fallback:
                    return f"Clima (web):\n{fallback}"
                return mensaje_error_amigable(
                    str(e), f"consultar el clima en {ciudad}"
                )

        if any(
            x in comando_lower
            for x in [
                "envia un whatsapp",
                "envía un whatsapp",
                "manda un whatsapp",
                "manda whatsapp",
                "send whatsapp",
            ]
        ):
            numero, mensaje = extraer_whatsapp(comando_original)
            if not numero or not mensaje:
                return "Usa: envia un whatsapp al NUMERO diciendo MENSAJE"
            try:
                send_whatsapp_message(numero, mensaje)
                return f"He preparado el mensaje de WhatsApp para {numero}."
            except Exception as e:
                return mensaje_error_amigable(str(e), "enviar WhatsApp")

        if any(
            x in comando_lower
            for x in [
                "envia un correo",
                "envía un correo",
                "manda un correo",
                "send an email",
            ]
        ):
            destinatario, asunto, mensaje = extraer_email(comando_original)
            if not destinatario or not asunto or not mensaje:
                return "Usa: envia un correo a DESTINATARIO asunto ASUNTO mensaje MENSAJE"
            enviado, error = send_email(destinatario, asunto, mensaje)
            if enviado:
                return f"He enviado el correo a {destinatario}."
            return f"No pude enviar el correo: {error}"

        if comando_lower.startswith(("ve a ", "ir a ", "navega a ")):
            destino = comando_original.split(" ", 2)[-1]
            return ir_a_url(destino)

        if comando_lower.startswith(("abre la pagina ", "abre la página ")):
            return ir_a_url(comando_original.split(" ", 3)[-1])

        if (
            comando_lower.startswith(("busca en ", "buscar en "))
            and "busca en youtube" not in comando_lower
            and "busca en google" not in comando_lower
        ):
            prefijo = (
                "busca en "
                if comando_lower.startswith("busca en ")
                else "buscar en "
            )
            resto = comando_original[len(prefijo) :]
            separadores = [" sobre ", " de ", " acerca de "]
            for separador in separadores:
                indice = resto.lower().find(separador)
                if indice != -1:
                    sitio = resto[:indice].strip()
                    consulta = resto[indice + len(separador) :].strip()
                    return buscar_en_sitio_web(sitio, consulta)

        if comando_lower.startswith(("abre ", "abrir ", "ábreme ", "abreme ", "ábrelo ", "ejecuta ", "lanza ")):
            for prefix in ["ábreme ", "abreme ", "ábrelo ", "abre ", "abrir ", "ejecuta ", "lanza "]:
                if comando_lower.startswith(prefix):
                    destino = comando_original[len(prefix):].strip()
                    break
            return abrir_aplicacion(destino)

        if comando_lower.startswith(("cierra ", "cerrar ", "ciérrame ", "cierrame ")):
            for prefix in ["ciérrame ", "cierrame ", "cierra ", "cerrar "]:
                if comando_lower.startswith(prefix):
                    destino = comando_original[len(prefix):].strip()
                    break
            return cerrar_aplicacion(destino)

        if comando_lower.startswith(("enfoca ", "activa ", "pon al frente ")):
            partes = comando_original.split(" ", 1)
            if len(partes) == 2:
                return activar_ventana(partes[1])

        if (
            "busca en youtube" in comando_lower
            or "tráiler de" in comando_lower
            or "trailer de" in comando_lower
            or "youtube" in comando_lower
        ):
            termino = extraer_consulta_busqueda(
                comando_original,
                [
                    "busca en youtube",
                    "youtube",
                    "muéstrame el tráiler de",
                    "muéstrame el trailer de",
                    "muestrame el tráiler de",
                    "muestrame el trailer de",
                    "tráiler de",
                    "trailer de",
                ],
            )
            if not termino:
                return "No entendí qué quieres buscar en YouTube."
            play_on_youtube(termino)
            return f"Abriendo YouTube para buscar: {termino}."

        if "busca en google" in comando_lower:
            termino = extraer_consulta_busqueda(
                comando_original, ["busca en google"]
            )
            if not termino:
                return "No entendí qué quieres buscar en Google."
            search_on_google(termino)
            return f"Buscando {termino} en Google."

        if comando_lower.startswith(("pega ", "pegar ")):
            partes = comando_original.split(" ", 1)
            if len(partes) == 2:
                return pegar_texto_en_ventana(partes[1])

        if comando_lower.startswith(("reemplaza con ", "reemplazar con ")):
            prefijo = (
                "reemplaza con "
                if comando_lower.startswith("reemplaza con ")
                else "reemplazar con "
            )
            return reemplazar_seleccion(comando_original[len(prefijo) :])

        if comando_lower in {"selecciona todo", "seleccionar todo"}:
            return seleccionar_todo()

        if comando_lower in {
            "nueva linea",
            "nueva línea",
            "salto de linea",
            "salto de línea",
            "enter",
        }:
            return nueva_linea()

        if comando_lower.startswith(
            (
                "nueva linea ",
                "nueva línea ",
                "tab ",
                "tabulacion ",
                "tabulación ",
            )
        ):
            numero = extraer_numero(comando_lower) or 1
            if comando_lower.startswith(("tab ", "tabulacion ", "tabulación ")):
                return tabular(numero)
            return nueva_linea(numero)

        if comando_lower in {
            "tab",
            "tabulacion",
            "tabulación",
            "indent",
            "indenta",
        }:
            return tabular()

        if any(
            x in comando_lower
            for x in [
                "spotify",
                "musica",
                "música",
                "cancion",
                "canción",
                "reproduccion",
                "reproducción",
                "video",
                "pelicula",
                "película",
            ]
        ):
            if any(
                x in comando_lower
                for x in [
                    "pausa",
                    "pausar",
                    "reanuda",
                    "reproduce",
                    "play",
                    "continua",
                    "inicia",
                    "comienza",
                    "empieza",
                ]
            ):
                return controlar_multimedia("play_pause")
            if any(
                x in comando_lower
                for x in [
                    "siguiente",
                    "siguiente pista",
                    "siguiente cancion",
                    "siguiente canción",
                    "next",
                ]
            ):
                return controlar_multimedia("next")
            if any(x in comando_lower for x in ["anterior", "previa", "previous"]):
                return controlar_multimedia("previous")
            if any(
                x in comando_lower
                for x in [
                    "deten",
                    "detener",
                    "stop",
                    "para la musica",
                    "para la música",
                    "termina",
                    "finaliza",
                    "acaba",
                ]
            ):
                return controlar_multimedia("stop")

        if any(
            x in comando_lower
            for x in [
                "play",
                "pausa",
                "reanuda",
                "siguiente cancion",
                "siguiente canción",
                "reproduce",
                "inicia",
                "comienza",
            ]
        ):
            if any(x in comando_lower for x in ["siguiente", "next"]):
                return controlar_multimedia("next")
            return controlar_multimedia("play_pause")

        if any(
            x in comando_lower
            for x in [
                "anterior",
                "pista anterior",
                "cancion anterior",
                "canción anterior",
            ]
        ):
            return controlar_multimedia("previous")

        if any(
            x in comando_lower
            for x in [
                "deten la musica",
                "detén la música",
                "stop multimedia",
                "deten el video",
                "detén el video",
                "para el video",
                "para la musica",
                "termina el video",
                "finaliza el video",
            ]
        ):
            return controlar_multimedia("stop")

        if "video" in comando_lower:
            if "pantalla completa" in comando_lower or "fullscreen" in comando_lower:
                return controlar_video("fullscreen")
            elif (
                "salir pantalla completa" in comando_lower
                or "exit fullscreen" in comando_lower
            ):
                return controlar_video("exit_fullscreen")
            elif any(
                x in comando_lower
                for x in ["play", "iniciar", "reproducir", "continuar"]
            ):
                return controlar_video("play")
            elif any(x in comando_lower for x in ["pausa", "pausar", "parar"]):
                return controlar_video("pause")

        if (
            "volumen" in comando_lower
            or "silencia" in comando_lower
            or "mute" in comando_lower
        ):
            numero = extraer_numero(comando_lower)
            if any(x in comando_lower for x in ["silencia", "mute", "mutear"]):
                return silenciar_audio(True)
            if any(
                x in comando_lower
                for x in [
                    "activa el audio",
                    "quita mute",
                    "quitar mute",
                    "desmutea",
                    "desmutear",
                    "reactiva sonido",
                ]
            ):
                return silenciar_audio(False)
            if (
                any(
                    x in comando_lower
                    for x in ["pon", "ajusta", "fija", "establece"]
                )
                and numero is not None
            ):
                return fijar_volumen(numero)
            if any(
                x in comando_lower
                for x in ["sube", "subir", "aumenta", "más", "mas"]
            ):
                return cambiar_volumen(numero if numero is not None else 10)
            if any(
                x in comando_lower
                for x in ["baja", "bajar", "reduce", "menos"]
            ):
                return cambiar_volumen(-(numero if numero is not None else 10))
            if (
                any(
                    x in comando_lower for x in ["cuanto", "qué", "que", "nivel"]
                )
                and "volumen" in comando_lower
            ):
                return f"El volumen actual está en {obtener_volumen_actual()}%."

        if any(
            x in comando_lower
            for x in [
                "apaga el ordenador",
                "apagar pc",
                "shutdown",
                "apaga la pc",
                "apaga el pc",
                "apagar el ordenador",
                "apagar ordenador",
                "apaga el equipo",
                "apagar equipo",
            ]
        ):
            if IS_WINDOWS:
                subprocess.run(["shutdown", "/s", "/t", "0"])
            else:
                subprocess.run(["shutdown", "-h", "now"])
            return "Apagando el ordenador."

        if any(
            x in comando_lower
            for x in [
                "reinicia el ordenador",
                "reiniciar pc",
                "restart",
                "reinicia la pc",
                "reinicia el pc",
                "reiniciar el ordenador",
                "reiniciar ordenador",
                "reinicia el equipo",
                "reiniciar equipo",
            ]
        ):
            if IS_WINDOWS:
                subprocess.run(["shutdown", "/r", "/t", "0"])
            else:
                subprocess.run(["reboot"])
            return "Reiniciando el ordenador."

        if (
            "bloquea la pantalla" in comando_lower
            or "lock screen" in comando_lower
        ):
            if IS_WINDOWS:
                subprocess.run(
                    ["rundll32.exe", "user32.dll,LockWorkStation"]
                )
            else:
                subprocess.run(["loginctl", "lock-session"])
            return "Pantalla bloqueada."

        if "hiberna" in comando_lower or "hibernate" in comando_lower:
            if IS_WINDOWS:
                subprocess.run(["shutdown", "/h"])
            else:
                subprocess.run(["systemctl", "hibernate"])
            return "Entrando en hibernación."

        if (
            "mueve el raton" in comando_lower
            or "mover raton" in comando_lower
            or "move mouse" in comando_lower
        ):
            nums = re.findall(r"\d+", comando_original)
            if len(nums) >= 2:
                x, y = int(nums[0]), int(nums[1])
                return mover_raton(x, y)
            return "Especifica coordenadas x y."

        if "click" in comando_lower:
            if "derecho" in comando_lower or "right" in comando_lower:
                return click_raton("right")
            elif "doble" in comando_lower or "double" in comando_lower:
                return doble_click()
            return click_raton("left")

        if "scroll" in comando_lower or "desplaza" in comando_lower:
            if "arriba" in comando_lower or "up" in comando_lower:
                return scroll_raton("arriba")
            return scroll_raton("abajo")

        if comando_lower.startswith(("ejecuta ", "run ")):
            cmd = comando_original.split(" ", 1)[1].strip()
            return ejecutar_comando_sistema(cmd)

        if any(
            x in comando_lower
            for x in ["ir a directorio", "ir a carpeta", "cambiar a", "cd"]
        ):
            partes = comando_original.split()
            path = partes[-1].strip() if partes else "."
            if "escritorio" in comando_lower:
                path = os.path.join(os.path.expanduser("~"), "Desktop")
            return cambiar_directorio(path)

        if "lista archivos" in comando_lower or "ls" == comando_lower.strip():
            path = "."
            if "escritorio" in comando_lower:
                path = os.path.join(os.path.expanduser("~"), "Desktop")
            return listar_archivos(path)

        if any(
            x in comando_lower
            for x in [
                "crea directorio",
                "crear directorio",
                "crea carpeta",
                "crear carpeta",
                "mkdir",
            ]
        ):
            name = comando_original.split()[-1].strip()
            path = name
            if "escritorio" in comando_lower:
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                path = os.path.join(desktop, name)
            return crear_directorio(path)

        if any(
            x in comando_lower
            for x in [
                "borra directorio",
                "borrar directorio",
                "borra carpeta",
                "borrar carpeta",
                "rmdir",
            ]
        ):
            name = comando_original.split()[-1].strip()
            path = name
            if "escritorio" in comando_lower:
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                path = os.path.join(desktop, name)
            return borrar_directorio(path)

        if "crea" in comando_lower and "archivo" in comando_lower:
            name = "archivo.txt"
            if "llamad" in comando_lower:
                idx = comando_original.lower().find("llamad")
                parts = comando_original[idx:].split()
                if len(parts) > 1:
                    name = parts[1]
                    if not name.endswith(".txt"):
                        name += ".txt"
            path = name
            if "escritorio" in comando_lower:
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                path = os.path.join(desktop, name)
            return crear_archivo(path, "")

        if (
            "lee la pantalla" in comando_lower
            or "que dice en pantalla" in comando_lower
            or "lee pantalla" in comando_lower
        ):
            return leer_pantalla()

        if any(
            x in comando_lower
            for x in [
                "como se llama este personaje",
                "quien es este personaje",
                "que personaje es este",
                "nombre del personaje",
            ]
        ):
            texto = leer_pantalla()
            if texto:
                prompt = (
                    f"El usuario pregunta por el nombre de un personaje en pantalla. "
                    f"Basado en este texto extraído: '{texto}', identifica el personaje."
                )
                return consultar_ia(prompt, cargar_memoria())
            return "No pude leer texto en la pantalla."

        if (
            "activa el inicio automatico" in comando_lower
            or "activar inicio automatico" in comando_lower
            or "inicio automatico" in comando_lower
        ):
            return activar_inicio_automatico()

        if (
            "desactiva el inicio automatico" in comando_lower
            or "desactivar inicio automatico" in comando_lower
        ):
            return desactivar_inicio_automatico()

    except Exception as e:
        return mensaje_error_amigable(str(e), "esa acción del sistema")

    if any(x in comando_lower for x in ["quien es", "quién es", "que es", "qué es", "define", "háblame de", "hablame de", "dime quien es", "dime quién es"]):
        termino = comando_lower
        for prefix in ["quien es", "quién es", "que es", "qué es", "define",
                       "háblame de", "hablame de", "dime quien es", "dime quién es"]:
            termino = termino.replace(prefix, "").strip()
        if termino:
            info_web = buscar_informacion_web(termino, max_results=5)
            prompt = (
                f"El usuario pregunta: '{comando_original}'\n\n"
                f"Información encontrada en internet y Wikipedia:\n{info_web}\n\n"
                f"Da una respuesta completa, precisa y detallada sobre '{termino}'. "
                f"Incluye datos relevantes: quién es, qué hace, datos importantes, "
                f"fechas, logros, nacionalidad, etc. Responde en español."
            )
            return consultar_ia_directa(prompt)

    # Enhanced web search fallback
    if any(
        x in comando_lower
        for x in [
            "buscar en internet",
            "busca en la red",
            "busca información",
            "busca informacion",
            "investiga",
            "averigua",
            "dime sobre",
            "háblame de",
            "hablame de",
            "información sobre",
            "informacion sobre",
        ]
    ):
        query = comando_lower
        for prefix in [
            "buscar en internet",
            "busca en la red",
            "busca información sobre",
            "busca informacion sobre",
            "investiga sobre",
            "averigua sobre",
            "dime sobre",
            "háblame de",
            "hablame de",
            "información sobre",
            "informacion sobre",
        ]:
            query = query.replace(prefix, "").strip()
        info = buscar_informacion_web(query)
        prompt = (
            f"El usuario pregunta: '{comando_original}'\n\n"
            f"Información encontrada en internet:\n{info}\n\n"
            f"Resume la información de forma clara y concisa en español."
        )
        return consultar_ia_directa(prompt)

    if comando_lower.startswith(("ejecuta como admin ", "ejecuta como administrador ")):
        prefijo = (
            "ejecuta como admin "
            if comando_lower.startswith("ejecuta como admin ")
            else "ejecuta como administrador "
        )
        return ejecutar_comando_sistema(
            comando_original[len(prefijo) :], como_admin=True
        )

    if comando_lower.startswith("ejecuta comando "):
        return ejecutar_comando_sistema(
            comando_original[len("ejecuta comando ") :]
        )

    if comando_lower.startswith("corre comando "):
        return ejecutar_comando_sistema(
            comando_original[len("corre comando ") :]
        )

    if comando_lower.startswith(("ejecuta powershell ", "ejecuta terminal ")):
        prefijo = (
            "ejecuta powershell "
            if comando_lower.startswith("ejecuta powershell ")
            else "ejecuta terminal "
        )
        return ejecutar_comando_sistema(comando_original[len(prefijo) :])

    if comando_lower.startswith("crea archivo "):
        resto = comando_original[len("crea archivo ") :]
        if " con contenido " in resto.lower():
            indice = resto.lower().find(" con contenido ")
            ruta = resto[:indice]
            contenido = resto[indice + len(" con contenido ") :]
            return crear_archivo(ruta, contenido)
        return crear_archivo(resto)

    if comando_lower.startswith("sobrescribe archivo "):
        resto = comando_original[len("sobrescribe archivo ") :]
        if " con " in resto.lower():
            indice = resto.lower().find(" con ")
            ruta = resto[:indice]
            contenido = resto[indice + len(" con ") :]
            return sobrescribir_archivo(ruta, contenido)
        return "Usa: sobrescribe archivo RUTA con CONTENIDO"

    if comando_lower.startswith("agrega al archivo "):
        resto = comando_original[len("agrega al archivo ") :]
        if " con " in resto.lower():
            indice = resto.lower().find(" con ")
            ruta = resto[:indice]
            contenido = resto[indice + len(" con ") :]
            return agregar_a_archivo(ruta, contenido)
        return "Usa: agrega al archivo RUTA con CONTENIDO"

    if comando_lower.startswith("lee archivo "):
        return leer_archivo(comando_original[len("lee archivo ") :])

    if comando_lower.startswith("lee documento "):
        ruta = comando_original[len("lee documento ") :]
        contenido, error = leer_documento(ruta)
        if error:
            return error
        return f"Contenido de {normalizar_ruta(ruta)}:\n{contenido}"

    if comando_lower.startswith("resume archivo ") or comando_lower.startswith(
        "resumen archivo "
    ):
        prefijo = (
            "resume archivo "
            if comando_lower.startswith("resume archivo ")
            else "resumen archivo "
        )
        return resumir_documento(comando_original[len(prefijo) :])

    if comando_lower.startswith("ejecuta instrucciones de archivo "):
        ruta = comando_original[len("ejecuta instrucciones de archivo ") :]
        return generar_plan_desde_documento(ruta)

    if comando_lower.startswith(
        "responde preguntas del archivo "
    ) or comando_lower.startswith("contesta preguntas del archivo "):
        prefijo = (
            "responde preguntas del archivo "
            if comando_lower.startswith("responde preguntas del archivo ")
            else "contesta preguntas del archivo "
        )
        ruta = comando_original[len(prefijo) :]
        return responder_preguntas_de_documento(ruta)

    if any(
        x in comando_lower
        for x in [
            "mira mi pantalla",
            "ver pantalla",
            "haz captura",
            "captura de pantalla",
        ]
    ):
        return capturar_pantalla()

    if "ventana activa" in comando_lower or "que ventana tengo" in comando_lower:
        return obtener_ventana_activa()

    if comando_lower.startswith("escribe en pantalla "):
        return escribir_en_ventana(
            comando_original[len("escribe en pantalla ") :]
        )

    if comando_lower.startswith("escribe codigo "):
        return escribir_codigo_en_editor(
            comando_original[len("escribe codigo ") :]
        )

    if comando_lower.startswith("pulsa "):
        return presionar_tecla(comando_original[len("pulsa ") :])

    if comando_lower.startswith("atajo "):
        return pulsar_teclas(comando_original[len("atajo ") :])

    if comando_lower in {"guarda archivo", "guardar archivo"}:
        return pulsar_teclas("ctrl+s")

    if comando_lower in {"copia", "copiar"}:
        return pulsar_teclas("ctrl+c")

    if comando_lower in {"corta", "cortar"}:
        return cortar()

    if comando_lower in {"pega", "pegar"}:
        return pulsar_teclas("ctrl+v")

    if comando_lower == "deshacer":
        return pulsar_teclas("ctrl+z")

    if comando_lower == "rehacer":
        return rehacer()

    if comando_lower in {
        "borra seleccion",
        "borra selección",
        "borrar seleccion",
        "borrar selección",
    }:
        return borrar_seleccion()

    return None


def consultar_ia_directa(prompt):
    """Query the AI directly without adding to memory or local action check."""
    try:
        cliente_groq = obtener_cliente_groq()
        if not cliente_groq:
            raise RuntimeError("Groq no configurado")
        completion = cliente_groq.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Eres Jarvis, asistente inteligente estilo Iron Man. "
                        "IMPORTANTE: Tu respuesta se lee en voz alta. Sé breve y conciso. "
                        "Responde en español con datos concretos y verificados. "
                        "Si hay información de internet en el prompt, úsala como fuente principal. "
                        "Nunca inventes datos sobre personas, celebridades o eventos. "
                        "Puedes hacer comparativas detalladas. Dirígete al usuario como 'señor'."
                    ),
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
        )
        respuesta = (completion.choices[0].message.content or "").strip()
        if respuesta:
            return dirigir_como_senor(respuesta)
    except Exception:
        pass
    if ollama is not None:
        try:
            response = ollama.chat(
                model="llama3.2:1b",
                messages=[{"role": "user", "content": prompt}],
            )
            respuesta = (response["message"]["content"] or "").strip()
            if respuesta:
                return dirigir_como_senor(respuesta)
        except Exception:
            pass
    return "No pude procesar esa consulta ahora mismo."


# === CEREBRO HIBRIDO ===
def consultar_ia(pregunta, historial):
    respuesta_directa = ejecutar_accion_local(pregunta)
    if respuesta_directa:
        return dirigir_como_senor(respuesta_directa)

    # Check if it needs web data
    pregunta_lower = pregunta.lower()
    necesita_web = any(
        x in pregunta_lower
        for x in [
            "precio",
            "cuanto cuesta",
            "cuánto cuesta",
            "mejor",
            "recomienda",
            "opinión",
            "opinion",
            "review",
            "reseña",
            "vs",
            "versus",
            "actual",
            "hoy",
            "ahora",
            "último",
            "ultimo",
            "nuevo",
            "nueva",
            "2024",
            "2025",
            "2026",
            "quien es",
            "quién es",
            "que es",
            "qué es",
            "cuantos años",
            "cuántos años",
            "donde nacio",
            "dónde nació",
            "de donde es",
            "de dónde es",
            "biografia",
            "biografía",
            "historia de",
            "cuanto gana",
            "cuánto gana",
            "cuanto mide",
            "cuánto mide",
            "cuanto pesa",
            "cuánto pesa",
            "cuantos goles",
            "cuántos goles",
            "palmares",
            "palmarés",
            "premios",
            "filmografia",
            "filmografía",
        ]
    )

    contexto_web = ""
    if necesita_web:
        contexto_web = buscar_informacion_web(pregunta)

    historial_modelo = preparar_contexto_modelo(historial)
    historial.append({"role": "user", "content": pregunta})

    contenido_pregunta = pregunta
    if contexto_web:
        contenido_pregunta = (
            f"{pregunta}\n\n[Datos de internet para tu respuesta]:\n{contexto_web}"
        )

    try:
        cliente_groq = obtener_cliente_groq()
        if not cliente_groq:
            raise RuntimeError("Groq no configurado")
        completion = cliente_groq.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=historial_modelo
            + [{"role": "user", "content": contenido_pregunta}],
            temperature=0.2,
        )
        respuesta = (completion.choices[0].message.content or "").strip()
        if not respuesta:
            raise RuntimeError("Respuesta vacía de Groq")
    except Exception:
        if ollama is not None:
            try:
                response = ollama.chat(
                    model="llama3.2:1b",
                    messages=historial_modelo
                    + [{"role": "user", "content": contenido_pregunta}],
                )
                respuesta = (response["message"]["content"] or "").strip()
                if not respuesta:
                    raise RuntimeError("Respuesta vacía de Ollama")
            except Exception:
                respuesta = "No pude responder con Groq ni con Ollama."
        else:
            respuesta = "No pude responder. Groq no disponible y Ollama no instalado."

    respuesta = dirigir_como_senor(respuesta)
    historial.append({"role": "assistant", "content": respuesta})
    guardar_memoria(historial)
    return respuesta


# === INTERFAZ GRAFICA - IRON MAN HUD ===
class JarvisApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("J.A.R.V.I.S. // STARK INDUSTRIES")
        self.geometry("1400x900")
        self.minsize(1200, 780)
        self.configure(fg_color="#010409")
        self.historial = cargar_memoria()
        self.escucha_activa = threading.Event()
        self.escucha_activa.set()
        self.ia_hablando = threading.Event()
        self.tts_lock = threading.Lock()
        self.tts_engine = None
        self.base_tts_engine = None
        self.voice_id = None
        self.reactor_mode = "idle"
        self.archivo_activo = None
        self.esperando_ruta_archivo = False
        self._reactor_angle = 0
        self._reactor_pulse = 0
        self._reactor_anim_id = None
        self._scan_offset = 0
        self._interrupt_monitor_active = threading.Event()
        self.modo_conversacion = threading.Event()
        self._conversation_timeout = 30
        self._last_interaction_time = 0
        self._interrupted_text = None
        self._ultima_respuesta_larga = None
        self._ultimo_comando = None
        self.RUTA_GUARDAR_TXT = os.path.join(
            "C:\\Users\\Javi\\Desktop\\jarvis\\cosas"
            if IS_WINDOWS
            else os.path.expanduser("~/jarvis/cosas")
        )

        try:
            engine = pyttsx3.init()
            self.voice_id = elegir_voz_masculina(engine)
            self.base_tts_engine = engine
            self.tts_engine = engine
            try:
                if self.voice_id:
                    engine.setProperty("voice", self.voice_id)
                engine.setProperty("rate", 178)
                engine.setProperty("volume", 1.0)
            except Exception:
                pass
        except Exception:
            self.voice_id = None
            self.base_tts_engine = None

        self._build_ui()
        self._start_reactor_animation()
        self.protocol("WM_DELETE_WINDOW", self.cerrar_app)
        self.iniciar_hilo_escucha()
        self.iniciar_hilo_agenda()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # === HEADER ===
        self.header_frame = ctk.CTkFrame(
            self,
            fg_color="#060d18",
            border_color="#00d4ff",
            border_width=2,
            corner_radius=12,
        )
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)

        self.title_block = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.title_block.grid(row=0, column=0, sticky="w", padx=16, pady=10)

        self.kicker = ctk.CTkLabel(
            self.title_block,
            text="STARK INDUSTRIES // AUTONOMOUS DEFENSE PLATFORM",
            font=("Consolas", 10, "bold"),
            text_color="#ff6b35",
        )
        self.kicker.pack(anchor="w")

        self.label = ctk.CTkLabel(
            self.title_block,
            text="J . A . R . V . I . S .",
            font=("Consolas", 36, "bold"),
            text_color="#00d4ff",
        )
        self.label.pack(anchor="w", pady=(2, 0))

        self.subtitle = ctk.CTkLabel(
            self.title_block,
            text="Just A Rather Very Intelligent System  //  v4.3.0-MARK-VII",
            font=("Consolas", 11),
            text_color="#5a8a9e",
        )
        self.subtitle.pack(anchor="w")

        self.header_right = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.header_right.grid(row=0, column=1, sticky="e", padx=16, pady=10)

        self.mode_badge = ctk.CTkLabel(
            self.header_right,
            text="VOICE CORE ONLINE",
            font=("Consolas", 11, "bold"),
            text_color="#010409",
            fg_color="#00d4ff",
            corner_radius=4,
            padx=10,
            pady=4,
        )
        self.mode_badge.pack(anchor="e", pady=(0, 6))

        self.listen_badge = ctk.CTkLabel(
            self.header_right,
            text="HOTWORD ARMED",
            font=("Consolas", 10, "bold"),
            text_color="#00ff88",
        )
        self.listen_badge.pack(anchor="e")

        self.time_label = ctk.CTkLabel(
            self.header_right,
            text="",
            font=("Consolas", 10),
            text_color="#5a8a9e",
        )
        self.time_label.pack(anchor="e", pady=(4, 0))
        self._update_clock()

        # === MAIN SHELL ===
        self.shell_frame = ctk.CTkFrame(
            self,
            fg_color="#020610",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=12,
        )
        self.shell_frame.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 6))
        self.shell_frame.grid_columnconfigure(0, weight=3)
        self.shell_frame.grid_columnconfigure(1, weight=2)
        self.shell_frame.grid_rowconfigure(0, weight=3)
        self.shell_frame.grid_rowconfigure(1, weight=2)

        # === ARC REACTOR ===
        self.reactor_frame = ctk.CTkFrame(
            self.shell_frame,
            fg_color="#010409",
            border_color="#00d4ff",
            border_width=2,
            corner_radius=12,
        )
        self.reactor_frame.grid(
            row=0, column=0, sticky="nsew", padx=(10, 5), pady=(10, 5)
        )
        self.reactor_frame.grid_columnconfigure(0, weight=1)
        self.reactor_frame.grid_rowconfigure(1, weight=1)

        self.reactor_header = ctk.CTkFrame(
            self.reactor_frame, fg_color="transparent"
        )
        self.reactor_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        self.reactor_header.grid_columnconfigure(0, weight=1)

        self.reactor_title = ctk.CTkLabel(
            self.reactor_header,
            text="ARC REACTOR // NEURAL CORE",
            font=("Consolas", 12, "bold"),
            text_color="#ff6b35",
        )
        self.reactor_title.grid(row=0, column=0, sticky="w")

        self.reactor_hint = ctk.CTkLabel(
            self.reactor_header,
            text='Wake: "Jarvis"',
            font=("Consolas", 10),
            text_color="#5a8a9e",
        )
        self.reactor_hint.grid(row=0, column=1, sticky="e")

        self.reactor_canvas = tk.Canvas(
            self.reactor_frame,
            width=640,
            height=360,
            bg="#010409",
            highlightthickness=0,
            bd=0,
        )
        self.reactor_canvas.grid(
            row=1, column=0, sticky="nsew", padx=8, pady=(0, 8)
        )
        self.reactor_canvas.bind(
            "<Configure>", lambda _: self._draw_reactor(self.reactor_mode)
        )

        # === JARVIS RESPONSE PANEL ===
        self.console_frame = ctk.CTkFrame(
            self.shell_frame,
            fg_color="#020610",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=12,
        )
        self.console_frame.grid(
            row=1, column=0, sticky="nsew", padx=(10, 5), pady=(0, 10)
        )
        self.console_frame.grid_columnconfigure(0, weight=1)
        self.console_frame.grid_rowconfigure(2, weight=1)

        self.console_top = ctk.CTkFrame(
            self.console_frame, fg_color="transparent"
        )
        self.console_top.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        self.console_top.grid_columnconfigure(0, weight=1)

        self.console_title = ctk.CTkLabel(
            self.console_top,
            text="JARVIS RESPONSE",
            font=("Consolas", 12, "bold"),
            text_color="#00d4ff",
        )
        self.console_title.grid(row=0, column=0, sticky="w")

        self.console_hint = ctk.CTkLabel(
            self.console_top,
            text="Voice Output Active",
            font=("Consolas", 10),
            text_color="#00ff88",
        )
        self.console_hint.grid(row=0, column=1, sticky="e")

        # User query display
        self.query_label = ctk.CTkLabel(
            self.console_frame,
            text="",
            font=("Consolas", 11),
            text_color="#5a8a9e",
            anchor="w",
            wraplength=550,
        )
        self.query_label.grid(row=1, column=0, sticky="ew", padx=14, pady=(2, 2))

        # Main response display
        self.response_textbox = ctk.CTkTextbox(
            self.console_frame,
            font=("Consolas", 13),
            fg_color="#010409",
            text_color="#b0e0ff",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=8,
            scrollbar_button_color="#0a3d5c",
            scrollbar_button_hover_color="#ff6b35",
            wrap="word",
        )
        self.response_textbox.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 6))
        self.response_textbox.insert(
            "end",
            "STARK OS v4.3.0 // ALL SYSTEMS NOMINAL\n\n"
            'Di "Jarvis" para iniciar conversación.\n'
            "Todas las respuestas se dan por voz.\n"
            "Si la respuesta es larga, recibirás un resumen.\n"
            'Di "guárdalo en archivo" para guardar la respuesta completa.',
        )

        # Hidden log textbox for internal logging
        self.textbox = ctk.CTkTextbox(
            self.console_frame,
            font=("Consolas", 10),
            fg_color="#010409",
            text_color="#5a8a9e",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=8,
            height=60,
            wrap="word",
        )
        self.textbox.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))

        # === SIDE PANEL ===
        self.side_frame = ctk.CTkFrame(
            self.shell_frame,
            fg_color="#040a14",
            border_color="#ff6b35",
            border_width=2,
            corner_radius=12,
        )
        self.side_frame.grid(
            row=0, column=1, rowspan=2, sticky="nsew", padx=(5, 10), pady=10
        )
        self.side_frame.grid_columnconfigure(0, weight=1)
        self.side_frame.grid_rowconfigure(4, weight=1)
        self.side_frame.grid_rowconfigure(6, weight=3)

        self.side_title = ctk.CTkLabel(
            self.side_frame,
            text="SYSTEM STATUS",
            font=("Consolas", 16, "bold"),
            text_color="#00d4ff",
        )
        self.side_title.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 2))

        self.side_subtitle = ctk.CTkLabel(
            self.side_frame,
            text="MARK VII INTERFACE",
            font=("Consolas", 9, "bold"),
            text_color="#ff6b35",
        )
        self.side_subtitle.grid(row=1, column=0, sticky="w", padx=12)

        self.status_label = ctk.CTkLabel(
            self.side_frame,
            text="Estado: esperando 'Jarvis'",
            font=("Consolas", 14, "bold"),
            text_color="#b0e0ff",
        )
        self.status_label.grid(row=2, column=0, sticky="w", padx=12, pady=(10, 4))

        self.orbit_status = ctk.CTkLabel(
            self.side_frame,
            text="IDLE",
            font=("Consolas", 10, "bold"),
            text_color="#00ff88",
        )
        self.orbit_status.grid(row=3, column=0, sticky="w", padx=12, pady=(0, 6))

        self.hud_metrics = ctk.CTkTextbox(
            self.side_frame,
            font=("Consolas", 10),
            fg_color="#010409",
            text_color="#7ab8d4",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=8,
            wrap="word",
            height=95,
        )
        self.hud_metrics.grid(row=4, column=0, sticky="ew", padx=12, pady=(0, 6))

        self.commands_title = ctk.CTkLabel(
            self.side_frame,
            text="CAPABILITIES",
            font=("Consolas", 10, "bold"),
            text_color="#ff6b35",
        )
        self.commands_title.grid(row=5, column=0, sticky="w", padx=12, pady=(0, 4))

        self.commands_box = ctk.CTkTextbox(
            self.side_frame,
            font=("Consolas", 10),
            fg_color="#010409",
            text_color="#7ab8d4",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=8,
            wrap="word",
        )
        self.commands_box.grid(row=6, column=0, sticky="nsew", padx=12, pady=(0, 8))

        self.btn_mute = ctk.CTkButton(
            self.side_frame,
            text="INTERRUPT VOICE",
            command=self.detener_voz,
            height=36,
            font=("Consolas", 11, "bold"),
            fg_color="#8b0000",
            hover_color="#b22222",
            text_color="#ffffff",
            corner_radius=6,
        )
        self.btn_mute.grid(row=7, column=0, sticky="ew", padx=12, pady=(0, 6))

        self.arc_label = ctk.CTkLabel(
            self.side_frame,
            text="ARC REACTOR STABLE",
            font=("Consolas", 9, "bold"),
            text_color="#00ff88",
        )
        self.arc_label.grid(row=8, column=0, sticky="w", padx=12, pady=(0, 12))

        # === FOOTER ===
        self.footer_frame = ctk.CTkFrame(
            self,
            fg_color="#060d18",
            border_color="#0a3d5c",
            border_width=1,
            corner_radius=10,
        )
        self.footer_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
        self.footer_frame.grid_columnconfigure(0, weight=0)
        self.footer_frame.grid_columnconfigure(1, weight=0)
        self.footer_frame.grid_columnconfigure(2, weight=1)

        self.footer_left = ctk.CTkLabel(
            self.footer_frame,
            text="STARK INDUSTRIES",
            font=("Consolas", 10, "bold"),
            text_color="#5a8a9e",
        )
        self.footer_left.grid(row=0, column=0, sticky="w", padx=12, pady=8)

        self.footer_right = ctk.CTkLabel(
            self.footer_frame,
            text="READY",
            font=("Consolas", 11, "bold"),
            text_color="#00ff88",
        )
        self.footer_right.grid(row=0, column=1, sticky="w", padx=(8, 12), pady=8)

        self.input_frame = ctk.CTkFrame(self.footer_frame, fg_color="transparent")
        self.input_frame.grid(row=0, column=2, sticky="ew", padx=(4, 10), pady=6)
        self.input_frame.grid_columnconfigure(0, weight=1)

        self.input_entry = ctk.CTkEntry(
            self.input_frame,
            placeholder_text="Escribe una orden para Jarvis...",
            font=("Consolas", 12),
            height=32,
            fg_color="#010409",
            border_color="#00d4ff",
            border_width=1,
            text_color="#b0e0ff",
            corner_radius=6,
        )
        self.input_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.input_entry.bind("<Return>", self.enviar_orden_texto_evento)

        self.send_button = ctk.CTkButton(
            self.input_frame,
            text="EXECUTE",
            width=90,
            command=self.enviar_orden_texto,
            fg_color="#00d4ff",
            hover_color="#00a5cc",
            text_color="#010409",
            font=("Consolas", 11, "bold"),
            corner_radius=6,
        )
        self.send_button.grid(row=0, column=1, sticky="e")

        self._draw_reactor("idle")
        self._render_hud_metrics("  LISTENER: HOTWORD READY\n")
        self._render_command_catalog()

    def _update_clock(self):
        now = datetime.now()
        self.time_label.configure(
            text=now.strftime("%H:%M:%S // %d-%m-%Y")
        )
        self.after(1000, self._update_clock)

    def log(self, mensaje):
        self.after(0, self._append_log, mensaje)

    def _draw_reactor(self, mode="idle"):
        self.reactor_mode = mode
        palette = {
            "idle": {
                "core": "#00d4ff",
                "glow": "#0088aa",
                "accent": "#005577",
                "ring": "#003344",
                "particle": "#00aacc",
                "text": "#00d4ff",
            },
            "listening": {
                "core": "#00ff88",
                "glow": "#00cc66",
                "accent": "#008844",
                "ring": "#004422",
                "particle": "#00ee77",
                "text": "#00ff88",
            },
            "speaking": {
                "core": "#ff6b35",
                "glow": "#cc5522",
                "accent": "#994411",
                "ring": "#662200",
                "particle": "#ff8855",
                "text": "#ff6b35",
            },
        }
        colors = palette.get(mode, palette["idle"])

        canvas = self.reactor_canvas
        canvas.delete("all")
        width = int(canvas.winfo_width() or 640)
        height = int(canvas.winfo_height() or 360)
        cx = width // 2
        cy = height // 2

        # HUD corner brackets
        bracket_len = 40
        bracket_color = colors["accent"]
        for bx, by, dx, dy in [
            (30, 30, 1, 1),
            (width - 30, 30, -1, 1),
            (30, height - 30, 1, -1),
            (width - 30, height - 30, -1, -1),
        ]:
            canvas.create_line(
                bx, by, bx + bracket_len * dx, by,
                fill=bracket_color, width=2,
            )
            canvas.create_line(
                bx, by, bx, by + bracket_len * dy,
                fill=bracket_color, width=2,
            )

        # Scan line
        scan_y = (self._scan_offset % height)
        canvas.create_line(
            0, scan_y, width, scan_y,
            fill=colors["ring"], width=1,
        )

        # Outer ring with rotation
        r_outer = min(width, height) * 0.38
        canvas.create_oval(
            cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer,
            outline=colors["ring"], width=2,
        )

        # Rotating arc segments
        angle = self._reactor_angle
        for i in range(6):
            start = angle + i * 60
            canvas.create_arc(
                cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer,
                start=start, extent=35,
                outline=colors["accent"], style="arc", width=3,
            )

        # Middle ring
        r_mid = r_outer * 0.75
        canvas.create_oval(
            cx - r_mid, cy - r_mid, cx + r_mid, cy + r_mid,
            outline=colors["glow"], width=3,
        )

        # Counter-rotating arcs
        for i in range(4):
            start = -angle * 1.5 + i * 90
            canvas.create_arc(
                cx - r_mid, cy - r_mid, cx + r_mid, cy + r_mid,
                start=start, extent=50,
                outline=colors["glow"], style="arc", width=2,
            )

        # Inner ring
        r_inner = r_outer * 0.52
        canvas.create_oval(
            cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner,
            outline=colors["core"], width=4,
        )

        # Spokes
        for i in range(8):
            a = (angle + i * 45) * math.pi / 180
            x1 = cx + int(r_inner * 0.95 * math.cos(a))
            y1 = cy + int(r_inner * 0.95 * math.sin(a))
            x2 = cx + int(r_mid * 1.05 * math.cos(a))
            y2 = cy + int(r_mid * 1.05 * math.sin(a))
            canvas.create_line(x1, y1, x2, y2, fill=colors["accent"], width=2)

        # Pulsing core
        pulse = abs(math.sin(self._reactor_pulse * 0.1)) * 0.3 + 0.7
        r_core = r_outer * 0.32 * pulse
        canvas.create_oval(
            cx - r_core, cy - r_core, cx + r_core, cy + r_core,
            fill=colors["ring"], outline=colors["core"], width=3,
        )

        r_core2 = r_core * 0.6
        canvas.create_oval(
            cx - r_core2, cy - r_core2, cx + r_core2, cy + r_core2,
            fill=colors["accent"], outline=colors["glow"], width=2,
        )

        r_core3 = r_core * 0.25
        canvas.create_oval(
            cx - r_core3, cy - r_core3, cx + r_core3, cy + r_core3,
            fill=colors["core"], outline="",
        )

        # HUD horizontal lines
        line_y_off = r_outer * 0.15
        canvas.create_line(
            cx - r_outer - 30, cy - line_y_off,
            cx - r_outer, cy - line_y_off,
            fill=colors["accent"], width=1,
        )
        canvas.create_line(
            cx + r_outer, cy - line_y_off,
            cx + r_outer + 30, cy - line_y_off,
            fill=colors["accent"], width=1,
        )
        canvas.create_line(
            cx - r_outer - 30, cy + line_y_off,
            cx - r_outer, cy + line_y_off,
            fill=colors["accent"], width=1,
        )
        canvas.create_line(
            cx + r_outer, cy + line_y_off,
            cx + r_outer + 30, cy + line_y_off,
            fill=colors["accent"], width=1,
        )

        # Status labels
        mode_labels = {
            "idle": "STANDBY",
            "listening": "RECEIVING",
            "speaking": "TRANSMITTING",
        }
        canvas.create_text(
            cx, cy + r_outer + 25,
            text=mode_labels.get(mode, "STANDBY"),
            fill=colors["text"],
            font=("Consolas", 10, "bold"),
        )
        canvas.create_text(
            cx, cy - r_outer - 20,
            text="J.A.R.V.I.S.",
            fill=colors["text"],
            font=("Consolas", 14, "bold"),
        )

        # Small data readouts
        canvas.create_text(
            50, height - 20,
            text=f"PWR: {int(pulse * 100)}%",
            fill=colors["accent"],
            font=("Consolas", 9),
            anchor="w",
        )
        canvas.create_text(
            width - 50, height - 20,
            text=f"SYNC: {datetime.now().strftime('%H:%M:%S')}",
            fill=colors["accent"],
            font=("Consolas", 9),
            anchor="e",
        )

    def _start_reactor_animation(self):
        self._reactor_angle += 2
        self._reactor_pulse += 1
        self._scan_offset += 2
        if self._reactor_angle >= 360:
            self._reactor_angle = 0
        self._draw_reactor(self.reactor_mode)
        self._reactor_anim_id = self.after(50, self._start_reactor_animation)

    def _render_hud_metrics(self, hud_estado):
        conv_mode = "ON" if self.modo_conversacion.is_set() else "OFF"
        self.hud_metrics.configure(state="normal")
        self.hud_metrics.delete("1.0", "end")
        self.hud_metrics.insert(
            "end",
            f"CORE STATUS\n"
            f"{hud_estado}"
            f"CONV MODE:    {conv_mode}\n"
            f"VOICE ENGINE: ONLINE\n"
            f"WEB SEARCH:   ACTIVE\n"
            f"MULTI-CMD:    READY\n"
            f"INTERRUPT:    ARMED\n"
            f"MULTI-MON:    LINKED\n",
        )
        self.hud_metrics.configure(state="disabled")

    def _render_command_catalog(self):
        self.commands_box.configure(state="normal")
        self.commands_box.delete("1.0", "end")
        self.commands_box.insert(
            "end",
            "CONVERSACION FLUIDA\n"
            "- Di 'Jarvis' una vez\n"
            "- Sigue hablando sin repetir\n"
            "- 'adiós' / 'nada más' = salir\n\n"
            "MULTI-COMANDO\n"
            "- abre Opera y Spotify\n"
            "- abre X, pon Y en 2a pantalla\n\n"
            "WEB & COMPARATIVAS\n"
            "- compara iPhone vs Samsung\n"
            "- investiga sobre...\n"
            "- qué es mejor X o Y\n\n"
            "SISTEMA\n"
            "- abre / cierra [app]\n"
            "- mueve [app] a 2a pantalla\n"
            "- volumen al 50\n\n"
            "MEDIA & WEB\n"
            "- clima en Madrid\n"
            "- top noticias\n"
            "- busca en YouTube [tema]\n"
            "- traduce 'frase' en idioma\n\n"
            "ARCHIVOS\n"
            '- accede al archivo "ruta"\n'
            "- lee / resume el archivo\n",
        )
        self.commands_box.configure(state="disabled")

    def _append_log(self, mensaje):
        self.textbox.insert("end", f"{mensaje}\n")
        self.textbox.see("end")

    def mostrar_respuesta(self, pregunta, respuesta):
        """Update the response panel with the latest Jarvis response."""
        def _actualizar():
            self.query_label.configure(text=f">> {pregunta}")
            self.response_textbox.delete("1.0", "end")
            self.response_textbox.insert("end", respuesta)
            self.response_textbox.see("1.0")
        self.after(0, _actualizar)

    def _resumir_para_voz(self, texto):
        """Summarize a long response for voice output using AI."""
        try:
            cliente_groq = obtener_cliente_groq()
            if not cliente_groq:
                raise RuntimeError("Groq no disponible")
            completion = cliente_groq.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "Eres Jarvis. Resume el siguiente texto en máximo 3 frases cortas "
                            "para ser leído en voz alta. Sé conciso y directo. "
                            "Al final añade: 'Si desea la información completa, dígame guárdalo en archivo'."
                        ),
                    },
                    {"role": "user", "content": texto},
                ],
                temperature=0.1,
            )
            resumen = (completion.choices[0].message.content or "").strip()
            if resumen:
                return dirigir_como_senor(resumen)
        except Exception:
            pass
        # Fallback: first 200 chars + prompt to save
        corte = texto[:200].rsplit(" ", 1)[0]
        return f"{corte}... Si desea la información completa, dígame guárdalo en archivo."

    def _guardar_respuesta_txt(self, contenido, tema="respuesta"):
        """Save a response to a TXT file in the configured save path."""
        os.makedirs(self.RUTA_GUARDAR_TXT, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_limpio = re.sub(r'[^\w\s-]', '', tema)[:50].strip().replace(' ', '_')
        nombre_archivo = f"jarvis_{nombre_limpio}_{timestamp}.txt"
        ruta_completa = os.path.join(self.RUTA_GUARDAR_TXT, nombre_archivo)
        with open(ruta_completa, "w", encoding="utf-8") as f:
            f.write(f"JARVIS - Respuesta generada el {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write(f"Consulta: {tema}\n")
            f.write("=" * 60 + "\n\n")
            f.write(contenido)
        return ruta_completa

    def actualizar_estado(self, mensaje):
        def _actualizar():
            self.status_label.configure(text=mensaje)
            mensaje_lower = mensaje.lower()
            en_conv = self.modo_conversacion.is_set()

            if "escuchando" in mensaje_lower or "conversación activa" in mensaje_lower:
                badge_text = "CONVERSATION" if en_conv else "MIC ACTIVE"
                self.listen_badge.configure(
                    text=badge_text, text_color="#00ff88"
                )
                hud_estado = "  LISTENER: CONVERSATION\n" if en_conv else "  LISTENER: CAPTURING\n"
                self.footer_right.configure(
                    text="CONV MODE" if en_conv else "LISTENING",
                    text_color="#00ff88",
                )
                self.orbit_status.configure(
                    text="DIALOGUE" if en_conv else "RECEIVING",
                    text_color="#00ff88",
                )
                self.reactor_mode = "listening"
            elif "respondiendo" in mensaje_lower:
                self.listen_badge.configure(
                    text="SPEAKING", text_color="#ff6b35"
                )
                hud_estado = "  LISTENER: SPEAKING\n"
                self.footer_right.configure(
                    text="RESPONDING", text_color="#ff6b35"
                )
                self.orbit_status.configure(
                    text="TRANSMITTING", text_color="#ff6b35"
                )
                self.reactor_mode = "speaking"
            else:
                if en_conv:
                    self.listen_badge.configure(
                        text="CONVERSATION", text_color="#00ff88"
                    )
                    hud_estado = "  LISTENER: CONVERSATION\n"
                    self.footer_right.configure(
                        text="CONV MODE", text_color="#00ff88"
                    )
                    self.orbit_status.configure(
                        text="DIALOGUE", text_color="#00ff88"
                    )
                    self.reactor_mode = "listening"
                else:
                    self.listen_badge.configure(
                        text="HOTWORD ARMED", text_color="#00ff88"
                    )
                    hud_estado = "  LISTENER: HOTWORD READY\n"
                    self.footer_right.configure(
                        text="READY", text_color="#00ff88"
                    )
                    self.orbit_status.configure(
                        text="IDLE", text_color="#00d4ff"
                    )
                    self.reactor_mode = "idle"

            self._render_hud_metrics(hud_estado)

        self.after(0, _actualizar)

    def iniciar_hilo_escucha(self):
        threading.Thread(target=self.bucle_escucha_continua, daemon=True).start()

    def iniciar_hilo_agenda(self):
        threading.Thread(target=self.bucle_agenda, daemon=True).start()

    def bucle_agenda(self):
        while self.escucha_activa.is_set():
            try:
                pendientes = obtener_items_vencidos_agenda()
                for item in pendientes:
                    tipo = item.get("tipo", "recordatorio")
                    mensaje = item.get("mensaje", "Tienes una alerta.")
                    aviso = f"{tipo.upper()}: {mensaje}"
                    self.log(f"Jarvis: {aviso}")
                    self.hablar_async(dirigir_como_senor(aviso))
            except Exception as e:
                self.log(f">>> Error en agenda: {e}")
            time.sleep(12)

    def hablar_async(self, texto):
        threading.Thread(target=self._hablar, args=(texto,), daemon=True).start()

    def _hablar(self, texto):
        if not texto or not texto.strip():
            return
        self.log(f">>> [VOZ] Iniciando TTS: {texto[:80]}...")
        with self.tts_lock:
            try:
                self.ia_hablando.set()
                self._interrupt_monitor_active.set()
                threading.Thread(
                    target=self._monitor_voice_interrupt, daemon=True
                ).start()
                self._hablar_pyttsx3(texto)
            except Exception as e:
                self.log(f">>> [VOZ] pyttsx3 falló: {e}")
                self._hablar_fallback(texto)
            finally:
                self.ia_hablando.clear()
                self._interrupt_monitor_active.clear()
                self.tts_engine = None

    def _hablar_pyttsx3(self, texto):
        """Primary TTS via pyttsx3 with retry logic."""
        for intento in range(2):
            try:
                engine = pyttsx3.init()
                if not self.voice_id:
                    self.voice_id = elegir_voz_masculina(engine)
                if self.voice_id:
                    try:
                        engine.setProperty("voice", self.voice_id)
                    except Exception:
                        pass
                engine.setProperty("rate", 175)
                engine.setProperty("volume", 1.0)
                self.tts_engine = engine
                engine.say(texto)
                engine.runAndWait()
                try:
                    engine.stop()
                except Exception:
                    pass
                self.log(">>> [VOZ] TTS completado.")
                return
            except RuntimeError as e:
                self.log(f">>> [VOZ] RuntimeError (intento {intento+1}): {e}")
                try:
                    engine.stop()
                except Exception:
                    pass
                self.tts_engine = None
                time.sleep(0.3)
            except Exception as e:
                self.log(f">>> [VOZ] Error pyttsx3 (intento {intento+1}): {e}")
                self.tts_engine = None
                time.sleep(0.3)
        raise RuntimeError("pyttsx3 failed after retries")

    def _hablar_fallback(self, texto):
        """Fallback TTS using system speech commands."""
        self.log(">>> [VOZ] Usando fallback del sistema...")
        try:
            if IS_WINDOWS:
                texto_ps = texto.replace("'", "''")
                ps_script = (
                    "Add-Type -AssemblyName System.Speech; "
                    "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                    f"$s.Rate = 1; $s.Speak('{texto_ps}')"
                )
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command", ps_script],
                    timeout=60, capture_output=True,
                )
            else:
                subprocess.run(
                    ["espeak", "-v", "es", texto],
                    timeout=60, capture_output=True,
                )
            self.log(">>> [VOZ] Fallback completado.")
        except Exception as e:
            self.log(f">>> [VOZ] Fallback también falló: {e}")

    def _monitor_voice_interrupt(self):
        """Monitors microphone while TTS is speaking; stops TTS and captures what user said."""
        try:
            recognizer = sr.Recognizer()
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.3)
                base_energy = recognizer.energy_threshold
                threshold = base_energy * 2.0

                while self._interrupt_monitor_active.is_set() and self.ia_hablando.is_set():
                    try:
                        audio = recognizer.listen(
                            source, timeout=0.5, phrase_time_limit=2
                        )
                        if audio.frame_data:
                            import audioop
                            rms = audioop.rms(audio.frame_data, 2)
                            if rms > threshold:
                                self.log(">>> Interrupción por voz detectada.")
                                self.detener_voz()
                                # Try to recognize what the user said during interruption
                                try:
                                    texto = self.reconocer_texto(recognizer, audio).strip()
                                    if texto:
                                        self._interrupted_text = texto
                                        self.log(f">>> Capturado durante interrupción: {texto}")
                                except Exception:
                                    pass
                                return
                    except sr.WaitTimeoutError:
                        continue
                    except Exception:
                        continue
        except Exception:
            pass

    def detener_voz(self):
        self._interrupt_monitor_active.clear()
        engine = self.tts_engine
        if engine:
            try:
                engine.stop()
                self.log(">>> Voz interrumpida.")
            except Exception:
                pass
        self.ia_hablando.clear()

    def contiene_wake_word(self, texto):
        texto_lower = texto.lower()
        words = texto_lower.split()
        for word in words:
            if difflib.get_close_matches(word, WAKE_WORDS, n=1, cutoff=0.52):
                return True
        return False

    def extraer_comando(self, texto):
        texto_limpio = texto.strip()
        texto_lower = texto_limpio.lower()
        words = texto_lower.split()
        for i, word in enumerate(words):
            matches = difflib.get_close_matches(word, WAKE_WORDS, n=1, cutoff=0.52)
            if matches:
                partes = texto_limpio.split()
                comando_partes = partes[:i] + partes[i + 1 :]
                comando = " ".join(comando_partes).strip()
                while comando and comando[0] in ",:;.-":
                    comando = comando[1:].strip()
                return comando
        return ""

    def limpiar_transcripcion(self, texto):
        texto = " ".join((texto or "").strip().split())
        if not texto:
            return ""

        reemplazos = {
            # Wake word corrections
            " jarviss": " jarvis",
            " yarvis": " jarvis",
            " jervis": " jarvis",
            " arvis": " jarvis",
            " jarbi ": " jarvis ",
            " yarbi ": " jarvis ",
            " harvis": " jarvis",
            " harvest": " jarvis",
            " javis": " jarvis",
            " jarves": " jarvis",
            " jarbis": " jarvis",
            " jarbs": " jarvis",
            " yarbis": " jarvis",
            " charvis": " jarvis",
            " jarvi ": " jarvis ",
            " garvis": " jarvis",
            " darvis": " jarvis",
            " tarvis": " jarvis",
            " jarviz": " jarvis",
            # Common Spanish misrecognitions
            " mostrarme ": " muéstrame ",
            " muestrame ": " muéstrame ",
            " trailer ": " tráiler ",
            " cancion ": " canción ",
            " reproduccion ": " reproducción ",
            " informacion ": " información ",
            " comparacion ": " comparación ",
            " volumen ": " volumen ",
            " aplicacion ": " aplicación ",
            " tambien ": " también ",
            " musica ": " música ",
            " pelicula ": " película ",
            " peliculas ": " películas ",
            " asi ": " así ",
            " pagina ": " página ",
            " abreme ": " abre ",
            " cierrame ": " cierra ",
            " ponme ": " pon ",
            # Common verb confusions
            " habre ": " abre ",
            " habreme ": " abre ",
            " habrir ": " abrir ",
            " ciérralo ": " cierra ",
            " cierrale ": " cierra ",
            # App name corrections
            " espotifai ": " spotify ",
            " espotify ": " spotify ",
            " spotifai ": " spotify ",
            " crome ": " chrome ",
            " cromo ": " chrome ",
            " gugle ": " google ",
            " guguel ": " google ",
            " discor ": " discord ",
            " discort ": " discord ",
            # Number corrections (common in Spanish speech)
            " diez ": " 10 ",
            " veinte ": " 20 ",
            " treinta ": " 30 ",
            " cuarenta ": " 40 ",
            " cincuenta ": " 50 ",
            " sesenta ": " 60 ",
            " setenta ": " 70 ",
            " ochenta ": " 80 ",
            " noventa ": " 90 ",
            " cien ": " 100 ",
            # Direction corrections
            " segunda pantaya ": " segunda pantalla ",
            " pantaya ": " pantalla ",
            " segund pantalla ": " segunda pantalla ",
        }
        texto_normalizado = f" {texto.lower()} "
        for origen, destino in reemplazos.items():
            texto_normalizado = texto_normalizado.replace(origen, destino)

        palabras = texto_normalizado.strip().split()
        corregidas = []
        for i, palabra in enumerate(palabras):
            if palabra == "1":
                anterior = palabras[i - 1] if i > 0 else ""
                siguiente = palabras[i + 1] if i + 1 < len(palabras) else ""
                if anterior not in {"xenoverse", "sinopsis"} and not siguiente.isdigit():
                    corregidas.append("el")
                    continue
            corregidas.append(palabra)

        return " ".join(corregidas).strip()

    def puntuar_transcripcion(self, texto):
        texto_limpio = self.limpiar_transcripcion(texto)
        if not texto_limpio:
            return -999

        palabras = texto_limpio.split()
        score = len(texto_limpio)
        score += len(palabras) * 3
        if self.contiene_wake_word(texto_limpio):
            score += 18
        if any(ch.isdigit() for ch in texto_limpio):
            score += 6
        if " 1 " in f" {texto_limpio} ":
            score -= 8
        return score

    def reconocer_texto(self, reconocedor, audio):
        try:
            detalle = reconocedor.recognize_google(
                audio, language="es-ES", show_all=True
            )
            alternativas = []
            if isinstance(detalle, dict):
                for alternativa in detalle.get("alternative", []):
                    transcripcion = alternativa.get("transcript", "").strip()
                    if transcripcion:
                        alternativas.append(transcripcion)
            if alternativas:
                mejor = max(alternativas, key=self.puntuar_transcripcion)
                return self.limpiar_transcripcion(mejor)
        except sr.UnknownValueError:
            return ""
        except Exception:
            pass

        try:
            texto = reconocedor.recognize_google(audio, language="es-ES")
            return self.limpiar_transcripcion(texto)
        except Exception:
            return ""

    def capturar_comando(self, reconocedor, source):
        self.actualizar_estado("Estado: escuchando orden")
        self.log(">>> Espero tu orden...")
        try:
            audio = reconocedor.listen(source, timeout=10, phrase_time_limit=30)
            texto = self.reconocer_texto(reconocedor, audio)
            comando = texto.strip()
            if comando:
                self.log(f"Tu: {comando}")
            return comando
        except sr.WaitTimeoutError:
            self.log(">>> Tiempo de espera agotado.")
            return ""
        except sr.UnknownValueError:
            self.log(">>> No entendí la orden.")
            return ""
        except Exception as e:
            self.log(f">>> Error capturando orden: {e}")
            return ""

    def procesar_comando(self, comando):
        if not comando:
            self.log(">>> No detecté una orden después de 'Jarvis'.")
            return

        comando_lower = comando.lower()

        # Check for "save to file" request
        if self._ultima_respuesta_larga and any(
            x in comando_lower
            for x in [
                "guárdalo", "guardalo", "guarda en archivo", "guárdalo en archivo",
                "crea el fichero", "crea un fichero", "crea el archivo",
                "crea un archivo", "guarda la información", "guarda la informacion",
                "ponlo en un archivo", "ponlo en un fichero",
                "guarda eso", "sí guárdalo", "si guardalo",
            ]
        ):
            ruta = self._guardar_respuesta_txt(
                self._ultima_respuesta_larga,
                self._ultimo_comando or "respuesta"
            )
            self._ultima_respuesta_larga = None
            respuesta = f"Información guardada en {ruta}"
            self.log(f"Jarvis: {respuesta}")
            self.mostrar_respuesta(comando, respuesta)
            self.actualizar_estado("Estado: respondiendo")
            self.hablar_async(respuesta)
            return

        # Check for multi-monitor commands
        if any(x in comando_lower for x in [
            "segunda pantalla", "segundo monitor",
            "pantalla secundaria", "monitor secundario",
            "otra pantalla", "otro monitor",
        ]) and any(x in comando_lower for x in ["pon", "mueve", "pasa", "lleva"]):
            app_name = _extract_app_from_command(comando)
            for noise in ["en la segunda pantalla", "a la segunda pantalla",
                          "al segundo monitor", "en el segundo monitor",
                          "a la pantalla secundaria", "en la pantalla secundaria"]:
                app_name = app_name.lower().replace(noise, "").strip()
            for prefix in ["ponmelo ", "ponme ", "ponlo ", "muévelo ", "muevelo ",
                           "pon ", "mueve ", "pasa ", "lleva "]:
                if app_name.startswith(prefix):
                    app_name = app_name[len(prefix):].strip()
            respuesta = mover_ventana_a_monitor(app_name, 1)
            self._responder_con_voz(comando, respuesta)
            return

        # Parse multi-commands
        comandos = separar_multi_comandos(comando)
        if len(comandos) > 1:
            self.log(f">>> Multi-comando detectado: {len(comandos)} órdenes")
            resultados = []
            for i, cmd in enumerate(comandos):
                self.log(f">>> Ejecutando [{i+1}/{len(comandos)}]: {cmd}")
                resultado = self._resolver_comando(cmd)
                resultados.append(resultado)
                time.sleep(0.5)
            resumen = ". ".join(resultados)
            self._responder_con_voz(comando, resumen)
            return

        self._ejecutar_comando_individual(comando)

    def _resolver_comando(self, comando):
        """Process a command and return the response text without side effects."""
        manejo_archivo = self.procesar_comando_archivo(comando)
        if manejo_archivo is not None:
            return manejo_archivo
        return consultar_ia(comando, self.historial)

    def _responder_con_voz(self, comando, respuesta):
        """Send response through voice, auto-summarizing if too long."""
        self.log(f"Jarvis: {respuesta}")
        self.mostrar_respuesta(comando, respuesta)
        self.actualizar_estado("Estado: respondiendo")

        if len(respuesta) > 300:
            self._ultima_respuesta_larga = respuesta
            self._ultimo_comando = comando
            resumen_voz = self._resumir_para_voz(respuesta)
            self.hablar_async(resumen_voz)
        else:
            self._ultima_respuesta_larga = None
            self.hablar_async(respuesta)

    def _ejecutar_comando_individual(self, comando):
        """Execute a single command with logging and voice output."""
        respuesta = self._resolver_comando(comando)
        self._responder_con_voz(comando, respuesta)
        return respuesta

    def extraer_ruta_de_comando(self, comando):
        comando = comando.strip()
        if not comando:
            return ""
        if '"' in comando:
            partes = comando.split('"')
            if len(partes) >= 3 and partes[1].strip():
                return partes[1].strip()
        if "'" in comando:
            partes = comando.split("'")
            if len(partes) >= 3 and partes[1].strip():
                return partes[1].strip()

        disparadores = [
            "accede al archivo ",
            "abre el archivo ",
            "abre archivo ",
            "usa el archivo ",
            "usa archivo ",
            "trabaja con el archivo ",
            "trabaja con archivo ",
            "lee archivo ",
            "lee documento ",
        ]
        texto = comando.lower()
        for d in disparadores:
            if d in texto:
                idx = texto.find(d)
                return comando[idx + len(d) :].strip()
        if (":\\" in comando or ":/" in comando) and "." in comando:
            return comando.strip().strip('"').strip("'")
        return ""

    def procesar_comando_archivo(self, comando):
        comando_original = comando.strip()
        comando_lower = comando_original.lower()

        if self.esperando_ruta_archivo:
            ruta = self.extraer_ruta_de_comando(comando_original)
            if not ruta:
                return "Necesito la ruta del archivo. Ejemplo: C:/Users/Javi/Desktop/notas.txt"
            ruta_norm = normalizar_ruta(ruta)
            if not os.path.exists(ruta_norm):
                return f"No encontré el archivo en {ruta_norm}."
            self.archivo_activo = ruta_norm
            self.esperando_ruta_archivo = False
            return (
                f"Archivo cargado: {ruta_norm}. "
                "Dime qué quieres que haga: leer, resumir, buscar, sobrescribir o agregar."
            )

        if (
            "archivo" in comando_lower
            and any(
                x in comando_lower
                for x in ["leer", "lea", "leas", "acceder", "abrir", "usar", "trabajar"]
            )
            and not self.extraer_ruta_de_comando(comando_original)
        ):
            self.esperando_ruta_archivo = True
            return "Pásame la ruta exacta del archivo."

        if any(
            x in comando_lower
            for x in [
                "accede al archivo",
                "abre el archivo",
                "abre archivo",
                "usa el archivo",
                "usa archivo",
                "trabaja con el archivo",
                "trabaja con archivo",
            ]
        ):
            ruta = self.extraer_ruta_de_comando(comando_original)
            if not ruta:
                self.esperando_ruta_archivo = True
                return 'Dime la ruta del archivo. Ejemplo: accede al archivo "C:/ruta/archivo.txt".'
            ruta_norm = normalizar_ruta(ruta)
            if not os.path.exists(ruta_norm):
                return f"No encontré el archivo en {ruta_norm}."
            self.archivo_activo = ruta_norm
            self.esperando_ruta_archivo = False
            return (
                f"Archivo cargado: {ruta_norm}. "
                "¿Qué quieres que haga? Leerlo, resumirlo, buscar texto, "
                "sobrescribir o agregar contenido."
            )

        if not self.archivo_activo:
            return None

        if any(
            x in comando_lower
            for x in [
                "cerrar archivo activo",
                "salir del archivo",
                "cancelar archivo",
            ]
        ):
            archivo = self.archivo_activo
            self.archivo_activo = None
            self.esperando_ruta_archivo = False
            return f"Archivo cerrado: {archivo}."

        if any(
            x in comando_lower
            for x in [
                "que pone",
                "qué pone",
                "lee el archivo",
                "leer archivo",
                "muestra contenido",
            ]
        ):
            contenido, error = leer_documento(self.archivo_activo)
            if error:
                return error
            return f"Contenido de {self.archivo_activo}:\n{contenido}"

        if any(x in comando_lower for x in ["resume", "resumen"]):
            return resumir_documento(self.archivo_activo)

        if any(
            x in comando_lower
            for x in [
                "ejecuta instrucciones",
                "haz lo que pone",
                "haz lo que dice",
            ]
        ):
            return generar_plan_desde_documento(self.archivo_activo)

        if any(
            x in comando_lower
            for x in [
                "responde las preguntas",
                "contesta las preguntas",
                "responde preguntas",
                "contesta preguntas",
            ]
        ):
            return responder_preguntas_de_documento(self.archivo_activo)

        if comando_lower.startswith("busca ") or "buscar " in comando_lower:
            consulta = (
                comando_original.lower()
                .replace("buscar ", "")
                .replace("busca ", "")
                .strip()
            )
            if not consulta:
                return "Dime qué texto quieres buscar dentro del archivo."
            contenido, error = leer_documento(self.archivo_activo, limite=30000)
            if error:
                return error
            idx = contenido.lower().find(consulta.lower())
            if idx == -1:
                return f"No encontré '{consulta}' en el archivo activo."
            inicio = max(0, idx - 220)
            fin = min(len(contenido), idx + len(consulta) + 220)
            fragmento = contenido[inicio:fin]
            return f"Encontré '{consulta}':\n...{fragmento}..."

        if comando_lower.startswith("sobrescribe con "):
            nuevo = comando_original[len("sobrescribe con ") :].strip()
            if not nuevo:
                return "Dime el contenido para sobrescribir."
            return sobrescribir_archivo(self.archivo_activo, nuevo)

        if comando_lower.startswith("agrega ") or comando_lower.startswith(
            "añade "
        ):
            pref = "agrega " if comando_lower.startswith("agrega ") else "añade "
            extra = comando_original[len(pref) :].strip()
            if not extra:
                return "Dime el contenido que quieres agregar."
            return agregar_a_archivo(self.archivo_activo, f"\n{extra}")

        contenido, error = leer_documento(self.archivo_activo)
        if error:
            return error
        instruccion = (
            f"Archivo activo: {self.archivo_activo}\n\n"
            f"Contenido:\n{contenido}\n\n"
            f"Instrucción: {comando_original}\n\n"
            "Responde de forma breve y accionable."
        )
        return consultar_ia(instruccion, self.historial)

    def enviar_orden_texto_evento(self, _event):
        self.enviar_orden_texto()

    def enviar_orden_texto(self):
        comando = self.input_entry.get().strip()
        if not comando:
            return
        self.input_entry.delete(0, "end")
        self.log(f"Tu (texto): {comando}")
        threading.Thread(
            target=self.procesar_comando, args=(comando,), daemon=True
        ).start()

    def _entrar_modo_conversacion(self):
        """Enter conversation mode — no wake word needed until timeout."""
        self.modo_conversacion.set()
        self._last_interaction_time = time.time()
        self.log(">>> Modo conversación ACTIVADO (no necesitas decir 'Jarvis')")
        self.actualizar_estado("Estado: modo conversación")

    def _salir_modo_conversacion(self):
        """Exit conversation mode — wake word required again."""
        if self.modo_conversacion.is_set():
            self.modo_conversacion.clear()
            self.log(">>> Modo conversación DESACTIVADO (di 'Jarvis' para hablar)")

    def _conversacion_activa(self):
        """Check if conversation mode is still active (not timed out)."""
        if not self.modo_conversacion.is_set():
            return False
        elapsed = time.time() - self._last_interaction_time
        if elapsed > self._conversation_timeout:
            self._salir_modo_conversacion()
            return False
        return True

    def _renovar_conversacion(self):
        """Refresh the conversation timeout."""
        self._last_interaction_time = time.time()

    def bucle_escucha_continua(self):
        reconocedor = sr.Recognizer()
        reconocedor.dynamic_energy_threshold = True
        reconocedor.pause_threshold = 0.9
        reconocedor.non_speaking_duration = 0.5
        reconocedor.phrase_threshold = 0.15
        reconocedor.operation_timeout = 15

        try:
            with sr.Microphone() as source:
                self.log(">>> Calibrando micrófono...")
                reconocedor.adjust_for_ambient_noise(source, duration=2)
                reconocedor.energy_threshold = max(
                    reconocedor.energy_threshold * 0.90, 150
                )
                self.log(
                    f">>> Umbral: {int(reconocedor.energy_threshold)} — Escucha activa."
                )

                while self.escucha_activa.is_set():
                    en_conversacion = self._conversacion_activa()

                    if en_conversacion:
                        self.actualizar_estado("Estado: conversación activa")
                    else:
                        self.actualizar_estado("Estado: esperando 'Jarvis'")

                    # Check if there's interrupted text to process
                    interrupted = self._interrupted_text
                    if interrupted:
                        self._interrupted_text = None
                        self.log(f">>> Procesando texto interrumpido: {interrupted}")
                        self._entrar_modo_conversacion()
                        comando = self.extraer_comando(interrupted) if self.contiene_wake_word(interrupted) else interrupted
                        if comando and len(comando.split()) > 1:
                            self._renovar_conversacion()
                            self.procesar_comando(comando)
                            continue

                    try:
                        # Wait for TTS to finish before listening (unless interrupted)
                        if self.ia_hablando.is_set():
                            time.sleep(0.2)
                            continue

                        timeout = 8 if en_conversacion else 1
                        audio = reconocedor.listen(
                            source, timeout=timeout, phrase_time_limit=30
                        )
                        texto = self.reconocer_texto(reconocedor, audio).strip()

                        if not texto:
                            continue

                        tiene_wake = self.contiene_wake_word(texto)

                        if en_conversacion:
                            # In conversation mode: process anything, no wake word needed
                            self._renovar_conversacion()

                            # Check for exit phrases
                            texto_lower = texto.lower()
                            if any(x in texto_lower for x in [
                                "hasta luego", "adiós", "adios", "nos vemos",
                                "eso es todo", "nada más", "nada mas",
                                "puedes descansar", "deja de escuchar",
                                "modo espera", "standby", "ya está",
                                "ya esta", "gracias jarvis", "vale gracias",
                                "ok gracias", "chao", "bye",
                                "descansa", "para ya", "cállate",
                                "callate", "silencio",
                            ]):
                                self._salir_modo_conversacion()
                                self.log("Jarvis: Entendido. Estaré aquí si me necesita.")
                                self.hablar_async("Entendido. Estaré aquí si me necesita.")
                                continue

                            # Extract command (remove wake word if present)
                            if tiene_wake:
                                comando = self.extraer_comando(texto)
                            else:
                                comando = texto

                            self.log(f"Tu: {comando}")
                            if comando and len(comando.split()) >= 1:
                                self.procesar_comando(comando)
                            continue

                        # Not in conversation mode: need wake word
                        if not tiene_wake:
                            continue

                        self.log(f">>> Wake word detectada: {texto}")

                        if self.ia_hablando.is_set():
                            self.log(">>> Interrumpiendo respuesta...")
                            self.detener_voz()
                            time.sleep(0.2)

                        # Enter conversation mode
                        self._entrar_modo_conversacion()

                        comando = self.extraer_comando(texto)
                        if not comando or len(comando.split()) <= 1:
                            self.log("Jarvis: Estoy aquí. Dígame.")
                            self.actualizar_estado("Estado: escuchando")
                            self.hablar_async("Estoy aquí. Dígame.")
                            try:
                                comando = self.capturar_comando(
                                    reconocedor, source
                                )
                            except Exception:
                                self.log(">>> No capté orden tras wake word.")
                                continue

                        if comando:
                            self._renovar_conversacion()
                            self.procesar_comando(comando)
                    except sr.WaitTimeoutError:
                        if en_conversacion:
                            # Timeout in conversation = exit conversation
                            self._salir_modo_conversacion()
                        continue
                    except sr.UnknownValueError:
                        continue
                    except sr.RequestError:
                        self.log(">>> Error de conexión con servicio de voz.")
                    except Exception as e:
                        self.log(f">>> Error en escucha: {e}")
        except OSError as e:
            self.log(f">>> No se pudo acceder al micrófono: {e}")
            self.log(">>> Modo solo texto activado.")

    def cerrar_app(self):
        self.escucha_activa.clear()
        self._interrupt_monitor_active.clear()
        if self._reactor_anim_id:
            self.after_cancel(self._reactor_anim_id)
        self.detener_voz()
        self.destroy()


if __name__ == "__main__":
    app = JarvisApp()
    app.mainloop()