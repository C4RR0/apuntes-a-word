"""
Gemini de Google con el plan GRATUITO de Google AI Studio (15/10/2026): calidad de
primera, rápido en cualquier ordenador y 0 €. Necesita internet y una clave gratis
(aistudio.google.com/apikey).

Hace lo mismo que Claude en la app (mismo proceso y mismas instrucciones): solo
cambia a quién se le pregunta. Elige solo el mejor modelo que ofrezca la clave
(primero los "pro", luego los "flash") y, si uno está saturado o sin cupo gratis,
pasa al siguiente.
"""
import json
import re
import ssl
import threading
import time
import urllib.error
import urllib.request

API = "https://generativelanguage.googleapis.com/v1beta"
EXCLUIR = ("image", "tts", "audio", "live", "embedding", "thinking", "exp", "learnlm", "gemma", "robotics",
           "computer-use", "nano", "omni")


class ClaveInvalida(Exception):
    pass


def limpiar_clave(texto):
    """Saca la clave de Google aunque se pegue con más texto alrededor."""
    m = re.search(r"AIza[0-9A-Za-z_\-]{30,}", texto or "")
    return m.group(0) if m else (texto or "").strip()


def _contexto_ssl():
    try:
        import certifi   # el Python de la app de Mac no ve los certificados del sistema
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _version(nombre):
    m = re.search(r"gemini-(\d+(?:\.\d+)?)", nombre)
    return float(m.group(1)) if m else 0.0


class Cliente:
    def __init__(self, clave, log=print):
        self.clave = limpiar_clave(clave)
        self.log = log
        self.ssl = _contexto_ssl()
        self.cerrojo = threading.Lock()
        self._modelos = None
        self.descartados = set()   # modelos sin cupo gratis para esta clave

    def _peticion(self, url, cuerpo=None, timeout=60):
        datos = json.dumps(cuerpo).encode("utf-8") if cuerpo is not None else None
        req = urllib.request.Request(url, datos, {"Content-Type": "application/json",
                                                  "x-goog-api-key": self.clave})
        with urllib.request.urlopen(req, timeout=timeout, context=self.ssl) as r:
            return json.load(r)

    def modelos(self):
        with self.cerrojo:
            if self._modelos is None:
                try:
                    datos = self._peticion(f"{API}/models?pageSize=1000")
                except urllib.error.HTTPError as e:
                    if e.code in (400, 401, 403):
                        raise ClaveInvalida("La clave de Google no es válida.") from e
                    raise
                nombres = [m["name"].removeprefix("models/") for m in datos.get("models", [])
                           if "generateContent" in m.get("supportedGenerationMethods", [])
                           and m["name"].removeprefix("models/").startswith("gemini")
                           and not any(x in m["name"] for x in EXCLUIR)]
                # pro antes que flash, flash antes que lite; dentro, la versión más nueva primero
                def orden(n):
                    familia = 0 if "pro" in n else 2 if "lite" in n else 1
                    return (familia, -_version(n), "preview" in n, len(n))
                self._modelos = sorted(nombres, key=orden)
                if not self._modelos:
                    raise RuntimeError("La clave de Google no ofrece ningún modelo Gemini.")
            return [m for m in self._modelos if m not in self.descartados]

    def generar(self, partes, sistema, max_tokens=65536):
        cuerpo = {
            "systemInstruction": {"parts": [{"text": sistema}]},
            "contents": [{"role": "user", "parts": partes}],
            "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.3},
        }
        ultimo = None
        for ronda in range(6):
            if ronda:
                espera = 30 * ronda
                self.log(f"   (Gemini está saturado; espero {espera} s y lo vuelvo a intentar)")
                time.sleep(espera)
            for modelo in self.modelos():
                try:
                    r = self._peticion(f"{API}/models/{modelo}:generateContent", cuerpo, timeout=900)
                except urllib.error.HTTPError as e:
                    texto = e.read().decode("utf-8", "replace")[:300]
                    ultimo = f"{modelo}: {e.code} {texto}"
                    if e.code in (401, 403) or (e.code == 400 and "API key" in texto):
                        raise ClaveInvalida("La clave de Google no es válida.") from e
                    if e.code == 429 and ("limit: 0" in texto or "free_tier" in texto and "PerDay" in texto):
                        self.descartados.add(modelo)   # sin cupo gratis (o agotado hoy): no insistir
                    continue
                except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
                    ultimo = f"{modelo}: {e}"
                    continue
                cand = (r.get("candidates") or [{}])[0]
                texto = "".join(p.get("text", "") for p in cand.get("content", {}).get("parts", [])
                                if not p.get("thought")).strip()
                if not texto:
                    ultimo = f"{modelo}: respuesta vacía ({cand.get('finishReason')})"
                    continue
                if cand.get("finishReason") == "MAX_TOKENS":
                    texto += "\n\n(El texto se ha cortado por ser demasiado largo.)"
                return texto, modelo
        raise RuntimeError(f"Gemini no responde ahora mismo; prueba dentro de un rato. ({ultimo})")
