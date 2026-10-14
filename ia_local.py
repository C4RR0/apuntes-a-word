"""IA gratis y sin internet para Apuntes a Word.

Usa Gemma 4 (Google, licencia abierta) con llama.cpp en el propio ordenador. La primera vez se
descargan el motor y el modelo (una sola vez); después todo funciona sin conexión y sin coste.

Para que no se salte nada:
  - cada diapositiva se lee dos veces: su texto exacto y su imagen (esquemas, rótulos, tablas);
  - se trabaja con pocas diapositivas a la vez;
  - después de cada respuesta se comprueba, frase a frase y cifra a cifra, que todo el contenido
    del original está en el texto generado; si falta algo se le pide que lo integre y, si aun así
    falta, se añade tal cual.
"""

import atexit
import base64
import io
import json
import os
import platform
import random
import re
import shutil
import socket
import subprocess
import sys
import tarfile
import time
import unicodedata
import urllib.request
import zipfile
from collections import Counter
from pathlib import Path

LLAMA_VERSION = "b11275"   # llama.cpp fijado (30/09/2026) para que no cambie solo
URL_LLAMA = "https://github.com/ggml-org/llama.cpp/releases/download/{v}/llama-{v}-bin-{plataforma}"
HF = "https://huggingface.co/unsloth/{repo}/resolve/main/{fichero}"
MODELOS = {
    # clave: (nombre visible, repo, gguf, bytes, mmproj, bytes)
    "maxima": ("Máxima calidad (Gemma 4 12B)", "gemma-4-12b-it-GGUF", "gemma-4-12b-it-Q4_K_M.gguf",
               7_121_861_440, "mmproj-F16.gguf", 175_115_840),
    "rapida": ("Rápida (Gemma 4 E4B)", "gemma-4-E4B-it-GGUF", "gemma-4-E4B-it-Q4_K_M.gguf",
               4_977_171_584, "mmproj-F16.gguf", 990_372_672),
}
if os.environ.get("APUNTES_MODELO_PRUEBA"):   # solo para la prueba automática del Mac: "repo|gguf|mmproj"
    _repo, _gguf, _mmproj = os.environ["APUNTES_MODELO_PRUEBA"].split("|")
    MODELOS["prueba"] = ("Modelo de prueba", _repo, _gguf, 0, _mmproj, 0)
CONTEXTO = 16384
PAGINAS_POR_GRUPO = 3
CARACTERES_POR_GRUPO = 2200


# ----------------------------------------------------------------------------- equipo y descargas

def memoria_gb():
    try:
        if sys.platform == "darwin":
            return int(subprocess.check_output(["sysctl", "-n", "hw.memsize"])) / 2**30
        if sys.platform == "win32":
            import ctypes

            class M(ctypes.Structure):
                _fields_ = [("l", ctypes.c_ulong), ("c", ctypes.c_ulong), ("total", ctypes.c_ulonglong),
                            ("a", ctypes.c_ulonglong * 6)]
            m = M()
            m.l = ctypes.sizeof(M)
            ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
            return m.total / 2**30
        return os.sysconf("SC_PAGE_SIZE") * os.sysconf("SC_PHYS_PAGES") / 2**30
    except Exception:
        return 8.0


def es_mac_apple():
    return sys.platform == "darwin" and platform.machine() == "arm64"


def modelo_recomendado():
    """En un Mac con chip Apple y 16 GB o más va bien el grande; si no, el rápido."""
    return "maxima" if es_mac_apple() and memoria_gb() >= 15 else "rapida"


def plataforma_llama():
    if sys.platform == "darwin":
        return "macos-arm64.tar.gz" if platform.machine() == "arm64" else "macos-x64.tar.gz"
    if sys.platform == "win32":
        return "win-vulkan-x64.zip"   # usa la tarjeta gráfica si la hay y, si no, el procesador
    return "ubuntu-x64.tar.gz"


def descargar(url, destino, tamaño, log, nombre):
    """Descarga con reanudación (si se corta, sigue donde iba) y comprobación del tamaño."""
    destino = Path(destino)
    if destino.exists() and (not tamaño or destino.stat().st_size == tamaño):
        return
    parcial = destino.with_suffix(destino.suffix + ".parcial")
    for intento in range(6):
        hecho = parcial.stat().st_size if parcial.exists() else 0
        peticion = urllib.request.Request(url, headers={"User-Agent": "ApuntesAWord"})
        if hecho:
            peticion.add_header("Range", f"bytes={hecho}-")
        try:
            with urllib.request.urlopen(peticion, timeout=60) as r:
                if hecho and r.status != 206:   # el servidor no reanuda: empezar de cero
                    hecho = 0
                total = tamaño or (hecho + int(r.headers.get("Content-Length", 0)))
                ultimo = -1
                with open(parcial, "ab" if hecho else "wb") as f:
                    while True:
                        bloque = r.read(1 << 20)
                        if not bloque:
                            break
                        f.write(bloque)
                        hecho += len(bloque)
                        if total:
                            pct = int(hecho * 100 / total)
                            if pct // 10 != ultimo:
                                ultimo = pct // 10
                                if total > 2**30:
                                    cuanto = f"{hecho / 2**30:.1f} de {total / 2**30:.1f} GB"
                                else:
                                    cuanto = f"{hecho / 2**20:.0f} de {total / 2**20:.0f} MB"
                                log(f"   Descargando {nombre}: {pct} % ({cuanto})")
            if not tamaño or parcial.stat().st_size == tamaño:
                parcial.replace(destino)
                return
        except Exception as e:
            log(f"   (la descarga se ha cortado: {e}; reintento)")
            time.sleep(5 * (intento + 1))
    raise RuntimeError(f"No he podido descargar {nombre}. Comprueba la conexión y vuelve a intentarlo.")


def ejecutable_servidor(carpeta):
    nombre = "llama-server.exe" if sys.platform == "win32" else "llama-server"
    encontrados = list(Path(carpeta).rglob(nombre))
    return encontrados[0] if encontrados else None


def preparar(carpeta, clave_modelo, log):
    """Descarga (solo la primera vez) el motor y el modelo. Devuelve (servidor, modelo, mmproj)."""
    carpeta = Path(carpeta)
    motor = carpeta / f"llama-{LLAMA_VERSION}"
    servidor = ejecutable_servidor(motor) if motor.exists() else None
    if servidor is None:
        log("   Primera vez: descargo el motor de la IA (unos 30 MB)...")
        motor.mkdir(parents=True, exist_ok=True)
        paquete = plataforma_llama()
        archivo = carpeta / paquete
        descargar(URL_LLAMA.format(v=LLAMA_VERSION, plataforma=paquete), archivo, 0, log, "el motor")
        if paquete.endswith(".zip"):
            with zipfile.ZipFile(archivo) as z:
                z.extractall(motor)
        else:
            with tarfile.open(archivo) as t:
                t.extractall(motor, filter="data")
        archivo.unlink(missing_ok=True)
        servidor = ejecutable_servidor(motor)
        if servidor is None:
            raise RuntimeError("El motor de la IA no se ha instalado bien.")
        if sys.platform != "win32":
            for f in servidor.parent.iterdir():
                if f.is_file() and not f.is_symlink():
                    f.chmod(0o755)
    _, repo, gguf, bytes_gguf, mmproj, bytes_mmproj = MODELOS[clave_modelo]
    modelo = carpeta / gguf
    vision = carpeta / f"{Path(gguf).stem}-{mmproj}"
    if not modelo.exists() or modelo.stat().st_size != bytes_gguf:
        log(f"   Primera vez: descargo la IA ({bytes_gguf / 2**30:.1f} GB). Solo se hace una vez; "
            "luego funciona sin internet.")
    descargar(HF.format(repo=repo, fichero=gguf), modelo, bytes_gguf, log, "la IA")
    descargar(HF.format(repo=repo, fichero=mmproj), vision, bytes_mmproj, log, "la vista de la IA")
    return servidor, modelo, vision


class Motor:
    """llama-server corriendo en este ordenador (solo escucha en 127.0.0.1)."""

    def __init__(self, servidor, modelo, vision, log):
        with socket.socket() as s:
            s.bind(("127.0.0.1", 0))
            self.puerto = s.getsockname()[1]
        self.log = log
        self.registro = open(Path(modelo).parent / "motor.log", "w", encoding="utf-8", errors="replace")
        opciones = {}
        if sys.platform == "win32":
            opciones["creationflags"] = 0x08000000   # sin ventana negra
        self.proceso = subprocess.Popen(
            [str(servidor), "-m", str(modelo), "--mmproj", str(vision), "--host", "127.0.0.1",
             "--port", str(self.puerto), "-c", str(CONTEXTO), "-np", "1", "--jinja", "--no-webui"],
            stdout=self.registro, stderr=subprocess.STDOUT, cwd=str(Path(servidor).parent), **opciones)
        atexit.register(self.cerrar)   # si se cierra la ventana a medias, la IA no se queda encendida
        inicio = time.time()
        while True:
            if self.proceso.poll() is not None:
                raise RuntimeError("La IA no ha podido arrancar (¿falta memoria? Prueba la opción «Rápida»). "
                                   f"Detalles en {self.registro.name}")
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{self.puerto}/health", timeout=5) as r:
                    if json.load(r).get("status") == "ok":
                        break
            except Exception:
                pass
            if time.time() - inicio > 600:
                self.cerrar()
                raise RuntimeError("La IA tarda demasiado en arrancar.")
            time.sleep(1)

    def cerrar(self):
        if self.proceso.poll() is None:
            self.proceso.terminate()
            try:
                self.proceso.wait(15)
            except subprocess.TimeoutExpired:
                self.proceso.kill()
        if not self.registro.closed:
            self.registro.close()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.cerrar()

    def pedir(self, contenido, sistema=None, max_tokens=4096,
              esquema=None, pensar=False, temperatura=0.2):
        mensajes = []
        if sistema:
            mensajes.append({"role": "system", "content": sistema})
        mensajes.append({"role": "user", "content": contenido})
        if os.environ.get("APUNTES_MAX_TOKENS"):   # solo en la prueba automática, para que sea rápida
            max_tokens = min(max_tokens, int(os.environ["APUNTES_MAX_TOKENS"]))
        cuerpo = {"messages": mensajes, "max_tokens": max_tokens, "temperature": temperatura,
                  "top_p": 0.95, "repeat_penalty": 1.05,
                  "chat_template_kwargs": {"enable_thinking": pensar}}
        if esquema:
            cuerpo["response_format"] = {"type": "json_schema", "json_schema": {"name": "r", "schema": esquema}}
        datos = json.dumps(cuerpo).encode("utf-8")
        for intento in range(3):
            try:
                peticion = urllib.request.Request(f"http://127.0.0.1:{self.puerto}/v1/chat/completions",
                                                  datos, {"Content-Type": "application/json"})
                with urllib.request.urlopen(peticion, timeout=3600) as r:
                    respuesta = json.load(r)
                eleccion = respuesta["choices"][0]
                texto = (eleccion["message"].get("content") or "").strip()
                if eleccion.get("finish_reason") == "length":
                    texto += "\n"
                return texto
            except Exception:
                if self.proceso.poll() is not None or intento == 2:
                    raise RuntimeError("La IA se ha detenido inesperadamente. "
                                       f"Detalles en {self.registro.name}")
                time.sleep(3)


# ----------------------------------------------------------------------------- comprobación de cobertura

VACIAS = set("""a al algo algun alguna algunas alguno algunos ante antes aqui asi aun cada como con contra cual
cuales cuando de del desde donde dos durante e el ella ellas ellos en entre era es esa esas ese eso esos esta
estan estas este esto estos fue fueron ha han hasta hay la las le les lo los mas me mediante mismo muy na ni no
nos o otra otras otro otros para pero poco por porque puede pueden que se segun ser si sin sobre son su sus
tambien tan tanto te tiene tienen todo todos tras tu un una unas uno unos y ya etc""".split())


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto.lower())
    return "".join(c for c in texto if not unicodedata.combining(c))


def raices(texto):
    """Conjunto de raíces de palabras con contenido (sin tildes ni palabras vacías)."""
    return {p[:6] for p in re.findall(r"[a-z0-9ñ]+", normalizar(texto))
            if p not in VACIAS and (len(p) > 2 or p.isdigit())}


def cifras(texto):
    return set(re.findall(r"\d+(?:[.,]\d+)?", texto.replace(" %", "%")))


def unidades(texto):
    """Trocea el texto original en ideas (líneas con al menos dos palabras con contenido)."""
    lineas, actual = [], ""
    for linea in texto.splitlines():
        s = linea.strip()
        if not s:
            continue
        nueva = bool(re.match(r"^([•▪■◦○●➢►✓\-–*]|\d+[.)]|[a-zA-Z][.)]\s)", s)) or not actual
        # una línea que no empieza por viñeta continúa la frase anterior si esta no estaba terminada
        if nueva or actual.rstrip().endswith((".", ":", ";")):
            if actual:
                lineas.append(actual)
            actual = s
        else:
            actual += " " + s
    if actual:
        lineas.append(actual)
    resultado = []
    for l in lineas:
        l = re.sub(r"^([•▪■◦○●➢►✓\-–*]\s*)", "", l).strip()
        if len(raices(l)) >= 2 or cifras(l):
            resultado.append(l)
    return resultado


def secuencia(texto):
    return [p[:6] for p in re.findall(r"[a-z0-9ñ]+", normalizar(texto))
            if p not in VACIAS and (len(p) > 2 or p.isdigit())]


def parejas(sec):
    """Parejas de palabras seguidas (o con una en medio): detectan frases que faltan aunque sus
    palabras sueltas salgan en otra parte."""
    return {(a, b) for i, a in enumerate(sec) for b in sec[i + 1:i + 3]}


def que_falta(original, generado, umbral=0.8, umbral_parejas=0.6):
    """Ideas y cifras del original que no aparecen en el texto generado.

    Calibrado con el Tema 2: 0 falsas alarmas con umbral_parejas=0.5 y 2 con 0.6 (que detecta
    bastantes más omisiones). Con umbral_parejas=0 solo quedan los casos claros."""
    sec = secuencia(generado)
    hay, hay_parejas = set(sec), parejas(sec)
    hay_cifras = cifras(generado)
    faltan = []
    for u in unidades(original):
        su = secuencia(u)
        r = set(su)
        cubiertas = sum(1 for x in r if x in hay) / max(len(r), 1)
        p = {(a, b) for a, b in zip(su, su[1:])}
        con_pareja = sum(1 for x in p if x in hay_parejas) / len(p) if p else 1
        sin_cifra = [c for c in cifras(u) if c not in hay_cifras]
        if cubiertas < umbral or con_pareja < umbral_parejas or sin_cifra:
            faltan.append(u)
    return faltan


# ----------------------------------------------------------------------------- lectura del PDF

def imagen_pagina(pdfium_doc, n, lado=896):
    pil = pdfium_doc[n - 1].render(scale=1.6).to_pil().convert("RGB")
    pil.thumbnail((lado, lado))
    return pil


def en_blanco(pil):
    gris = pil.convert("L").resize((64, 64))
    return max(gris.getdata()) - min(gris.getdata()) < 12


def a_base64(pil):
    buf = io.BytesIO()
    pil.save(buf, format="JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def es_mayusculas(s):
    letras = [c for c in s if c.isalpha()]
    return len(letras) >= 3 and sum(c.isupper() for c in letras) / len(letras) > 0.8


def leer_paginas(lector):
    """Texto de cada página sin las cabeceras y pies que se repiten en todas las diapositivas."""
    paginas = []
    for p in lector.pages:
        try:
            paginas.append(p.extract_text() or "")
        except Exception:
            paginas.append("")
    cuenta = Counter()
    for t in paginas:
        for l in {x.strip() for x in t.splitlines() if x.strip()}:
            cuenta[l] += 1
    limite = max(3, int(len(paginas) * 0.08))
    # distingue mayúsculas: la portada «PLACENTA» no es la cabecera «Placenta» de cada diapositiva
    repetidas = {l for l, n in cuenta.items() if n >= limite}
    limpias = []
    for t in paginas:
        lineas = [x.rstrip() for x in t.splitlines()
                  if x.strip() and x.strip() not in repetidas
                  and not re.fullmatch(r"\s*\d{1,3}\s*", x)]
        limpias.append("\n".join(lineas))
    return limpias, repetidas


def portada_de_apartado(texto):
    """Si la página es la portada de un apartado (p. ej. «PLACENTA»), devuelve su título."""
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    if not lineas or len(lineas) > 3 or len(texto) > 140:
        return None
    titulo = " ".join(l for l in lineas if es_mayusculas(l))
    if not titulo or len(titulo) < 4:
        return None
    titulo = re.sub(r"^\d+\s*[.)-]\s*", "", titulo).strip(" .:")
    return titulo.upper()


def planificar(paginas, n_paginas):
    """Lista de bloques: ("apartado", título) o ("grupo", [números de página])."""
    plan, grupo, chars = [], [], 0
    for n in range(1, n_paginas + 1):
        texto = paginas[n - 1]
        titulo = portada_de_apartado(texto) if n > 1 else None
        if n == 1 and len(texto) < 200:   # portada del tema: ya va como título del documento
            continue
        if titulo:
            if grupo:
                plan.append(("grupo", grupo))
                grupo, chars = [], 0
            plan.append(("apartado", titulo))
            continue
        if grupo and (len(grupo) >= PAGINAS_POR_GRUPO or chars + len(texto) > CARACTERES_POR_GRUPO):
            plan.append(("grupo", grupo))
            grupo, chars = [], 0
        grupo.append(n)
        chars += len(texto)
    if grupo:
        plan.append(("grupo", grupo))
    return plan


# ----------------------------------------------------------------------------- limpieza de la respuesta

LATEX = {r"\geq": "≥", r"\ge": "≥", r"\leq": "≤", r"\le": "≤", r"\rightarrow": "→", r"\to": "→",
         r"\leftarrow": "←", r"\uparrow": "↑", r"\downarrow": "↓", r"\times": "×", r"\pm": "±",
         r"\approx": "≈", r"\neq": "≠", r"\cdot": "·", r"\%": "%", r"\degree": "°", r"^\circ": "°",
         r"\circ": "°"}
SUBINDICES = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")


def sin_latex(texto):
    """La IA a veces escribe fórmulas en LaTeX ($\\geq$, $\\text{O}_2$): se pasan a texto normal."""
    def convertir(m):
        t = re.sub(r"\\(?:text|mathrm|mathbf)\{([^}]*)\}", r"\1", m.group(1))
        for k in sorted(LATEX, key=len, reverse=True):
            t = t.replace(k, LATEX[k])
        t = re.sub(r"_\{?(\d+)\}?", lambda x: x.group(1).translate(SUBINDICES), t)
        return t.replace("{", "").replace("}", "").replace("\\", "").strip()
    return re.sub(r"\$\$?([^$\n]{1,80}?)\$\$?", convertir, texto)


def ordenar_markdown(md, cabeceras=()):
    """Deja las respuestas en el formato exacto que entiende el conversor a Word."""
    md = sin_latex(re.sub(r"^```(?:markdown|md)?\s*|\s*```$", "", md.strip()))
    prohibidos = {" ".join(re.findall(r"[a-z0-9ñ]+", normalizar(c))) for c in cabeceras}
    lineas = md.splitlines()
    sangrias = [len(l) - len(l.lstrip(" ")) for l in lineas if re.match(r"^\s+[-*+•]\s", l)]
    unidad = min([s for s in sangrias if s > 0], default=2)
    unidad = 2 if unidad < 2 else min(unidad, 4)
    salida = []
    for l in lineas:
        l = l.replace("\t", "    ").rstrip()
        if re.fullmatch(r"\s*([-*_]\s*){3,}", l):
            continue
        if re.search(r"(?i)\(?\b(p[aá]gina|diapositiva|slide)\s*\d+\)?", l) and len(raices(l)) <= 2:
            continue
        l = re.sub(r"(?i)\s*\((p[aá]gina|diapositiva)s?\s*[\d\-–, y]+\)", "", l)
        # rellenos sin contenido que a veces escribe la IA: «(No se especifican detalles)»
        relleno = r"(no se (especifica|detalla|menciona|indica|describe)n?|sin (más )?detalles|no hay (más )?detalles)"
        l = re.sub(rf"(?i)\s*\(\s*{relleno}[^)]*\)", "", l)
        if re.match(rf"(?i)^\s*([-*+•]\s*)?(\*\*[^*]+\*\*\s*)?{relleno}", l):
            continue
        m = re.match(r"^(\s*)[-*+•]\s+(.*)", l)
        if m:
            nivel = min(len(m.group(1)) // unidad, 2)
            l = "  " * nivel + "- " + m.group(2)
        m = re.match(r"^(#{1,6})\s*(.*?)\s*#*$", l)
        if m:
            texto = m.group(2).replace("**", "").strip(" :")
            h = " ".join(re.findall(r"[a-z0-9ñ]+", normalizar(texto)))
            if any(h == p or (len(h) > 12 and h in p) for p in prohibidos):   # cabecera, no título
                continue
            l = f"{m.group(1)} {texto}"
        salida.append(l)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(salida)).strip()


def quitar_titulos_repetidos(md):
    salida, previo = [], {}
    for l in md.splitlines():
        m = re.match(r"^(#{2,4})\s+(.+)$", l)
        if m:
            nivel, clave = len(m.group(1)), normalizar(m.group(2)).strip(" .:")
            if any(previo.get(n) == clave for n in range(2, nivel + 1)):   # repite su título o el del apartado
                continue
            previo[nivel] = clave
            for mayor in list(previo):
                if mayor > nivel:
                    del previo[mayor]
        salida.append(l)
    return "\n".join(salida)


# ----------------------------------------------------------------------------- apuntes pasados a limpio

SISTEMA = (
    "Eres un profesor experto que prepara apuntes de estudio para una estudiante universitaria de "
    "Enfermería. Escribes en español de España con ortografía y tildes perfectas. Eres totalmente fiel "
    "al material: no inventas nada y no te saltas nada. Respondes solo con Markdown, sin introducciones, "
    "comentarios ni despedidas."
)

REGLAS_LIMPIO = """Pasa a limpio estas diapositivas como apuntes de estudio completos y bien redactados.
Reglas obligatorias:
1. Incluye TODO el contenido: cada frase, dato, cifra, porcentaje, clasificación (con todos sus tipos), definición, causa, signo, síntoma, complicación, tratamiento y cuidado de enfermería. Nada puede quedar fuera.
2. Lee también las imágenes de las diapositivas: si un esquema, tabla, rótulo o dibujo contiene información que no está en el texto extraído, intégrala en los apuntes. No describas fotos («la imagen muestra...»): escribe directamente la información útil. Ignora logotipos y adornos.
3. Usa el texto extraído para escribir los términos exactamente igual que en el original.
4. Une las frases cortadas por el formato de diapositiva. Mantén el orden del original.
5. Estructura: "### TÍTULO EN MAYÚSCULAS" para cada tema nuevo (suele ser el título de la diapositiva) y "#### Título" para títulos menores (tipos, clasificaciones, fórmulas). No repitas un título si el tema continúa.
6. Listas con "- " (subniveles con 2 espacios de sangría, como mucho 3 niveles). Cuando haya un término y su explicación: "- **Término:** explicación". Párrafos solo para introducciones breves.
7. Tablas en Markdown (| col | col |) cuando la información sea tabular. Sin celdas vacías.
8. Datos para memorizar de un vistazo (cifras clave, valores normales, fórmulas, definiciones esenciales) en líneas que empiecen por "> ", una idea por línea.
9. Lo MUY importante, que seguro cae en el examen, entre == == (como mucho una o dos veces).
10. Conceptos clave en **negrita** (solo los conceptos, no frases enteras).
11. No menciones números de página ni de diapositiva. No uses LaTeX ni el símbolo $: escribe ≥, ≤, →, O₂, CO₂...
12. No te limites a enumerar los rótulos de un dibujo: escribe lo que enseña (qué es, qué ocurre, qué consecuencias tiene), sin añadir nada que no se vea o no se lea en la diapositiva."""

REGLAS_IMAGENES = """13. Imágenes del PDF disponibles: {lista}. Coloca cada una que ilustre algo útil en su propia línea, justo después del texto al que corresponde, con el formato [[IDENTIFICADOR|pie de foto breve y descriptivo]], por ejemplo [[IMG_13_1|Placenta con el cordón umbilical]]."""

PROMPT_COMPLETAR = """Este es el texto original de unas diapositivas:
<original>
{original}
</original>

Y estos son los apuntes pasados a limpio que has escrito:
<apuntes>
{apuntes}
</apuntes>

Comparando con el original, en los apuntes FALTAN estas ideas o cifras (o están incompletas):
{faltan}

Devuelve los apuntes COMPLETOS ya corregidos, integrando cada una de ellas en su sitio, con el mismo formato y sin quitar nada de lo que ya estaba (conserva también los marcadores [[IMG_...|...]])."""


def _contexto(apartado, sub):
    partes = []
    if apartado:
        partes.append(f"Estas diapositivas pertenecen al apartado «{apartado}» (su título ya está puesto; no lo repitas).")
    if sub:
        partes.append(f"El último subapartado que se escribió fue «{sub}»: si estas diapositivas lo continúan, sigue sin repetir ese título.")
    return ("\n" + "\n".join(partes)) if partes else ""


def pasar_a_limpio_local(motor, lector, pdfium_doc, imagenes, log, titulo_doc=""):
    paginas, repetidas = leer_paginas(lector)
    cabeceras = [l for l in repetidas if len(raices(l)) >= 2] + ([titulo_doc] if titulo_doc else [])
    if paginas and len(paginas[0]) < 200:   # la portada del tema tampoco es un título interno
        cabeceras += [" ".join(paginas[0].split())] + [l.strip() for l in paginas[0].splitlines() if len(l) > 15]
    aviso = ""
    if cabeceras:
        aviso = ("\nEstas frases son cabeceras que se repiten en las diapositivas; no las uses como títulos: "
                 + "; ".join(f"«{c}»" for c in cabeceras[:8]) + ".")
    total = len(paginas)
    plan = planificar(paginas, total)
    hay_portadas = any(tipo == "apartado" for tipo, _ in plan)
    salida, apartado, sub = [], None, None
    inicio = time.time()
    hechas = 0
    for tipo, valor in plan:
        if tipo == "apartado":
            apartado, sub = valor, None
            salida.append(f"## {valor}")
            continue
        grupo = valor
        log(f"   Diapositivas {grupo[0]}-{grupo[-1]} de {total}..." + _quedan(inicio, hechas, total))
        contenido, original = [], []
        for n in grupo:
            pil = imagen_pagina(pdfium_doc, n)
            texto = paginas[n - 1].strip()
            if not texto and en_blanco(pil):
                continue
            original.append(texto)
            vacio = "(sin texto: la información está en la imagen)"
            contenido.append({"type": "text", "text": f"Diapositiva {n}. Texto extraído:\n{texto or vacio}\nImagen de la diapositiva {n}:"})
            contenido.append({"type": "image_url", "image_url": {"url": a_base64(pil)}})
        hechas += len(grupo)
        if not contenido:
            continue
        ids = [i for i in imagenes if int(i.split("_")[1]) in grupo]
        reglas = REGLAS_LIMPIO
        if ids:
            reglas += "\n" + REGLAS_IMAGENES.format(
                lista=", ".join(f"{i} (diapositiva {i.split('_')[1]})" for i in ids))
        contenido.append({"type": "text", "text": reglas + aviso + _contexto(apartado, sub)})
        md = ordenar_markdown(motor.pedir(contenido, SISTEMA, max_tokens=6000), cabeceras)
        texto_original = "\n".join(original)
        faltan = que_falta(texto_original, md)
        if faltan:
            log(f"   · Repaso: faltaban {len(faltan)} detalle(s); los integro")
            prompt = PROMPT_COMPLETAR.format(original=texto_original, apuntes=md,
                                             faltan="\n".join(f"- {f}" for f in faltan))
            nuevo = ordenar_markdown(motor.pedir(prompt, SISTEMA, max_tokens=7000), cabeceras)
            if len(nuevo) >= len(md) * 0.9:   # nunca aceptar una versión que recorte
                md = nuevo
            faltan = que_falta(texto_original, md, umbral_parejas=0)
            if faltan:   # garantía final: lo que siga faltando claramente se añade tal cual
                md += "\n" + "\n".join(f"- {f}" for f in faltan)
        for i in ids:   # imágenes que la IA no ha colocado: al final de su grupo
            if f"[[{i}" not in md:
                md += f"\n[[{i}|]]"
        titulos = re.findall(r"(?m)^###\s+(.+)$", md)
        if titulos:
            sub = titulos[-1]
        salida.append(md)
    md = quitar_titulos_repetidos("\n\n".join(salida))
    if not hay_portadas:   # PDF sin portadas de apartado: los temas de diapositiva pasan a apartados
        md = re.sub(r"(?m)^###(#?)\s", lambda m: "##" + m.group(1) + " ", md)
    elif not md.lstrip().startswith("## "):
        md = "## INTRODUCCIÓN\n\n" + md
    return md


def _quedan(inicio, hechas, total):
    if hechas < 3:
        return ""
    restante = (time.time() - inicio) / hechas * (total - hechas)
    return f" (faltan unos {max(1, round(restante / 60))} min para los apuntes)"


# ----------------------------------------------------------------------------- material de estudio

def apartados(md):
    """[(título, contenido)] de los apartados "## ", partiendo los muy largos por subapartados."""
    md = re.sub(r"\[\[IMG_[^\]]*\]\]", "", md)
    resultado = []
    for b in re.split(r"(?m)^##\s+", md):
        if not b.strip():
            continue
        titulo, _, cuerpo = b.partition("\n")
        titulo = re.sub(r"^\d+\.\s*", "", titulo).strip()
        if not cuerpo.strip():
            continue
        trozos, actual = [], ""
        for s in re.split(r"(?m)^(?=###\s)", cuerpo):
            if actual.strip() and len(actual) + len(s) > 9000:
                trozos.append(actual)
                actual = ""
            actual += s
        if actual.strip():
            trozos.append(actual)
        resultado += [(titulo, t.strip()) for t in trozos]
    return resultado


def _claves(md):
    """Lo que un resumen o esquema no puede olvidar: subapartados, términos en negrita y cifras."""
    titulos = [t.strip() for t in re.findall(r"(?m)^#{3,4}\s+(.+)$", md)]
    negritas = list(dict.fromkeys(t.strip(" :") for t in re.findall(r"\*\*(.+?)\*\*", md) if len(t) > 2))
    return titulos, negritas, cifras(md)


def _olvidos(original, generado, con_cifras=True):
    titulos, negritas, numeros = _claves(original)
    hay, hay_cifras = raices(generado), cifras(generado)

    def esta(t):
        r = raices(t)
        return not r or sum(x in hay for x in r) / len(r) >= 0.5

    faltan = [("subapartado", t) for t in titulos if not esta(t)]
    faltan += [("concepto", t) for t in negritas if not esta(t)]
    if con_cifras:
        faltan += [("cifra", c) for c in sorted(numeros - hay_cifras)]
    return faltan


def _lineas_con(original, olvidos):
    """Líneas del original con lo olvidado (para añadirlas si la IA no lo integra)."""
    lineas = []
    for tipo, buscado in olvidos:
        if tipo == "subapartado":
            continue
        for l in original.splitlines():
            s = l.strip()
            if buscado in s and s not in lineas and not s.startswith("#"):
                lineas.append(s)
                break
    return lineas


def resumen_local(motor, titulo, md, tipo):
    if tipo == "resumen":
        orden = ("Escribe el RESUMEN de este apartado para estudiar.\n"
                 "- Aproximadamente un tercio de la extensión original.\n"
                 "- Cubre TODOS los subapartados, en el mismo orden, cada uno con su título como \"#### Título\".\n"
                 "- Conserva, condensado, todo lo que se puede preguntar en un examen: todas las cifras y valores, "
                 "todas las clasificaciones nombrando todos sus tipos, definiciones esenciales, causas, signos, "
                 "complicaciones y cuidados de enfermería.\n"
                 "- Frases breves en viñetas \"- \"; conceptos clave en **negrita**; los datos para memorizar en "
                 "líneas que empiecen por \"> \".\n"
                 "- Condensa: agrupa en una sola línea, separadas por comas, las listas de elementos cortos.\n"
                 "- No repitas información ni añadas notas del tipo «no se especifica».\n"
                 "- No pongas el título del apartado ni frases introductorias.")
    else:
        orden = ("Haz el ESQUEMA de este apartado.\n"
                 "- Listas anidadas con \"- \" y 2 espacios de sangría por nivel (máximo 3 niveles).\n"
                 "- Palabras clave, no frases largas; conceptos en **negrita**.\n"
                 "- Incluye TODOS los subapartados en el mismo orden y todas las clasificaciones con todos sus tipos.\n"
                 "- No repitas información ni añadas notas del tipo «no se especifica».\n"
                 "- No pongas el título del apartado ni frases introductorias.")
    prompt = f"<apuntes apartado=\"{titulo}\">\n{md}\n</apuntes>\n\n{orden}"
    texto = ordenar_markdown(motor.pedir(prompt, SISTEMA, max_tokens=4000))
    olvidos = _olvidos(md, texto, con_cifras=(tipo == "resumen"))
    if olvidos:
        nombres = {"subapartado": "el subapartado", "concepto": "el concepto", "cifra": "la cifra"}
        arreglo = (prompt + "\n\nYa lo habías escrito así:\n<borrador>\n" + texto + "\n</borrador>\n\n"
                   "Pero te has dejado esto:\n" + "\n".join(f"- {nombres[t]} «{x}»" for t, x in olvidos)
                   + "\n\nDevuélvelo COMPLETO incluyendo cada uno en su sitio, con el mismo formato.")
        nuevo = ordenar_markdown(motor.pedir(arreglo, SISTEMA, max_tokens=4500))
        if len(nuevo) >= len(texto) * 0.9:
            texto = nuevo
        if tipo == "resumen":   # garantía final: lo que siga faltando se añade desde los apuntes
            restantes = _lineas_con(md, _olvidos(md, texto))
            if restantes:
                texto += "\n" + "\n".join(l if l.startswith((">", "-", "|")) else f"- {l}" for l in restantes)
    texto = re.sub(r"(?m)^#{2,3}\s", "#### ", texto)
    return texto


ESQUEMA_TEST = {
    "type": "object",
    "properties": {"preguntas": {"type": "array", "items": {
        "type": "object",
        "properties": {
            "pregunta": {"type": "string"},
            "correcta": {"type": "string"},
            "incorrectas": {"type": "array", "items": {"type": "string"}, "minItems": 3, "maxItems": 3},
            "explicacion": {"type": "string"},
        },
        "required": ["pregunta", "correcta", "incorrectas", "explicacion"]}}},
    "required": ["preguntas"],
}
ESQUEMA_COMPROBAR = {
    "type": "object",
    "properties": {"letra": {"type": "string", "enum": ["a", "b", "c", "d", "ninguna"]}},
    "required": ["letra"],
}
ESQUEMA_FICHAS = {
    "type": "object",
    "properties": {"fichas": {"type": "array", "items": {
        "type": "object",
        "properties": {"pregunta": {"type": "string"}, "respuesta": {"type": "string"}},
        "required": ["pregunta", "respuesta"]}}},
    "required": ["fichas"],
}


def _json(texto):
    try:
        return json.loads(texto)
    except Exception:
        m = re.search(r"\{.*\}", texto, re.S)
        try:
            return json.loads(m.group(0)) if m else {}
        except Exception:
            return {}


def test_local(motor, titulo, md, azar):
    n = max(2, min(8, round(len(md) / 1100)))
    prompt = (f"<apuntes apartado=\"{titulo}\">\n{md}\n</apuntes>\n\n"
              f"Crea {n} preguntas tipo test de examen sobre estos apuntes que, entre todas, cubran todos sus "
              "subapartados. Cada pregunta: enunciado claro, la respuesta correcta (respaldada literalmente "
              "por los apuntes), 3 respuestas incorrectas verosímiles pero inequívocamente falsas según los "
              "apuntes, y una explicación breve de por qué la correcta lo es. Las cuatro opciones deben tener "
              "una longitud y un estilo parecidos, para que la correcta no se note por ser la más larga. "
              "Evita «todas las anteriores» y «ninguna de las anteriores». Responde en JSON.")
    datos = _json(motor.pedir(prompt, SISTEMA, max_tokens=4000, esquema=ESQUEMA_TEST, temperatura=0.4))
    buenas = []
    for p in datos.get("preguntas", []):
        correcta = p.get("correcta", "").strip()
        opciones = [correcta] + [x.strip() for x in p.get("incorrectas", [])]
        if len(opciones) != 4 or len({normalizar(o) for o in opciones}) != 4 or not all(opciones):
            continue
        azar.shuffle(opciones)
        letra = "abcd"[opciones.index(correcta)]
        enunciado = p.get("pregunta", "").strip()
        texto = enunciado + "\n" + "\n".join(f"{l}) {o}" for l, o in zip("abcd", opciones))
        # segunda opinión: otra pasada responde la pregunta sin saber la solución
        verif = _json(motor.pedir(
            f"<apuntes>\n{md}\n</apuntes>\n\nResponde a esta pregunta tipo test usando SOLO los apuntes. "
            "Si hay más de una opción correcta o ninguna lo es, responde «ninguna».\n\n" + texto,
            SISTEMA, max_tokens=50, esquema=ESQUEMA_COMPROBAR, temperatura=0.0))
        if enunciado and verif.get("letra") == letra:
            buenas.append((texto, letra, correcta, p.get("explicacion", "").strip()))
    return buenas


def fichas_local(motor, titulo, md):
    n = max(2, min(10, round(len(md) / 900)))
    prompt = (f"<apuntes apartado=\"{titulo}\">\n{md}\n</apuntes>\n\n"
              f"Crea {n} fichas de repaso (pregunta y respuesta breve y exacta) que, entre todas, cubran los "
              "conceptos, cifras y clasificaciones más importantes de todos los subapartados. Todo tiene que "
              "estar respaldado por los apuntes. Responde en JSON.")
    datos = _json(motor.pedir(prompt, SISTEMA, max_tokens=3000, esquema=ESQUEMA_FICHAS, temperatura=0.3))
    return [(f.get("pregunta", "").strip(), f.get("respuesta", "").strip()) for f in datos.get("fichas", [])
            if f.get("pregunta", "").strip() and f.get("respuesta", "").strip()]


def renumerar_bloques(md):
    contador = iter(range(1, 10_000))
    return re.sub(r"(?m)^###\s+(?:\d+\.\s*)?(.+)$", lambda m: f"### {next(contador)}. {m.group(1)}", md)


def material_estudio_local(motor, apuntes_md, opciones, log):
    partes = apartados(apuntes_md)
    resumen, esquema, preguntas, fichas = [], [], [], []
    azar = random.Random(7)
    previo = None
    for i, (titulo, md) in enumerate(partes, 1):
        log(f"   Resumen, esquema, test y fichas: {titulo.capitalize()} ({i} de {len(partes)})...")
        cabecera = [f"### {titulo}"] if titulo != previo else []
        previo = titulo
        if opciones.get("resumen"):
            resumen += cabecera + [resumen_local(motor, titulo, md, "resumen")]
        if opciones.get("esquema"):
            esquema += cabecera + [resumen_local(motor, titulo, md, "esquema")]
        if opciones.get("test"):
            preguntas += test_local(motor, titulo, md, azar)
        if opciones.get("fichas"):
            fichas += fichas_local(motor, titulo, md)
    secciones = []
    if resumen:
        secciones.append("## Resumen\n" + renumerar_bloques("\n\n".join(resumen)))
    if esquema:
        secciones.append("## Esquema\n" + renumerar_bloques("\n\n".join(esquema)))
    if preguntas:
        cuerpo = "\n\n".join(f"{n}. {t}" for n, (t, _, _, _) in enumerate(preguntas, 1))
        soluciones = "\n".join(f"{n}. {l}) {o.rstrip('.')}. {x}" for n, (_, l, o, x) in enumerate(preguntas, 1))
        secciones.append(f"## Preguntas tipo test\n{cuerpo}\n\n### Soluciones\n{soluciones}")
    if fichas:
        secciones.append("## Fichas de repaso\n" + "\n\n".join(f"**P:** {p}\n**R:** {r}" for p, r in fichas))
    return sin_latex("\n\n".join(secciones))
