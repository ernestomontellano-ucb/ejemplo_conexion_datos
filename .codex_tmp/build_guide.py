from pathlib import Path
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path("\\\\wsl" + chr(36) + "\\Ubuntu\\home\\ernesto\\taller\\ejemplo_conexion_datos")
OUT = ROOT / "Guia_construccion_app_evento.docx"
APP = (ROOT / "assets/js/app.js").read_text(encoding="utf-8")
EVENTO = (ROOT / "assets/js/evento.js").read_text(encoding="utf-8")

def clean(text):
    return text.replace("§", chr(96)).replace("¤", chr(36))

def shade(target, fill):
    pr = target._tc.get_or_add_tcPr() if hasattr(target, "_tc") else target._p.get_or_add_pPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), fill)
    pr.append(node)

def cell_format(cell):
    pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "6")
        node.set(qn("w:color"), "D9D9D9")
        borders.append(node)
    pr.append(borders)
    margins = OxmlElement("w:tcMar")
    for name, value in (("top", 120), ("start", 140), ("bottom", 120), ("end", 140)):
        node = OxmlElement("w:" + name)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    pr.append(margins)

def heading(text, level=1):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.keep_with_next = True

def paragraph(text):
    doc.add_paragraph(text)

def bullet(text):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.space_after = Pt(3)

def numbered(text):
    p = doc.add_paragraph(text, style="List Number")
    p.paragraph_format.space_after = Pt(4)

def code(text):
    p = doc.add_paragraph(style="Code")
    p.add_run(clean(text).rstrip())
    shade(p, "F2F2F2")
    pr = p._p.get_or_add_pPr()
    border = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement("w:" + edge)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:color"), "D9D9D9")
        node.set(qn("w:space"), "5")
        border.append(node)
    pr.append(border)

def table(headers, rows, widths):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, value in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = value
        c.width = Inches(widths[i])
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(c, "252525")
        cell_format(c)
        for run in c.paragraphs[0].runs:
            run.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)
    for row_index, row in enumerate(rows):
        cells = t.add_row().cells
        for i, value in enumerate(row):
            c = cells[i]
            c.text = str(value)
            c.width = Inches(widths[i])
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell_format(c)
            if row_index % 2:
                shade(c, "F7F7F7")
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(0)
                for run in p.runs:
                    run.font.size = Pt(9.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

page_calls = 0

def page():
    global page_calls
    if page_calls == 0:
        doc.add_page_break()
    page_calls += 1

def page_number(p):
    run = p.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
    sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t"); text.text = "1"
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, sep, text, end])

doc = Document()
section = doc.sections[0]
section.page_width = Inches(8.5)
section.page_height = Inches(11)
section.top_margin = Inches(0.75)
section.bottom_margin = Inches(0.7)
section.left_margin = Inches(0.82)
section.right_margin = Inches(0.82)

normal = doc.styles["Normal"]
normal.font.name = "Aptos"
normal.font.size = Pt(11)
normal.font.color.rgb = RGBColor(32, 32, 32)
normal.paragraph_format.space_after = Pt(7)
normal.paragraph_format.line_spacing = 1.12

title_style = doc.styles["Title"]
title_style.font.name = "Aptos Display"
title_style.font.size = Pt(30)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor(0, 0, 0)

subtitle = doc.styles["Subtitle"]
subtitle.font.name = "Aptos"
subtitle.font.size = Pt(14)
subtitle.font.color.rgb = RGBColor(65, 65, 65)

for name, size in (("Heading 1", 21), ("Heading 2", 15), ("Heading 3", 12)):
    style = doc.styles[name]
    style.font.name = "Aptos Display"
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = Pt(15)
    style.paragraph_format.space_after = Pt(7)
    style.paragraph_format.keep_with_next = True

code_style = doc.styles.add_style("Code", 1)
code_style.font.name = "Consolas"
code_style.font.size = Pt(8.5)
code_style.font.color.rgb = RGBColor(30, 30, 30)
code_style.paragraph_format.left_indent = Inches(0.12)
code_style.paragraph_format.right_indent = Inches(0.12)
code_style.paragraph_format.space_before = Pt(5)
code_style.paragraph_format.space_after = Pt(8)
code_style.paragraph_format.line_spacing = 1

footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer.add_run("Guía para construir app js y evento js   |   ")
run.font.size = Pt(8.5)
run.font.color.rgb = RGBColor(110, 110, 110)
page_number(footer)

doc.add_paragraph().paragraph_format.space_after = Pt(38)
doc.add_paragraph("Guía para construir app js y evento js", style="Title")
doc.add_paragraph("Carga de datos JSON creación de tarjetas y navegación a una página de detalle", style="Subtitle")
doc.add_paragraph().paragraph_format.space_after = Pt(20)
p = doc.add_paragraph()
p.add_run("Propósito de la clase  ").bold = True
p.add_run("Construir dos archivos JavaScript que leen una agenda desde JSON muestran sus eventos y cargan el detalle correcto mediante un identificador incluido en la URL.")
table(["Archivo", "Responsabilidad"], [
    ["app.js", "Carga todos los eventos y construye las tarjetas"],
    ["evento.js", "Lee el id busca un evento y completa su ficha"],
    ["eventos.json", "Organiza categorías subcategorías y eventos"],
], [1.35, 5.15])
p = doc.add_paragraph()
p.add_run("Resultado esperado  ").bold = True
p.add_run("El estudiante podrá explicar las promesas el recorrido de datos anidados la generación de HTML y la selección de un evento concreto.")

page()
heading("Objetivos de aprendizaje")
for item in [
    "Distinguir una función normal una función asíncrona y una arrow function",
    "Identificar las promesas creadas por fetch respuesta json y una función async",
    "Recorrer categorías subcategorías y eventos con for of",
    "Elegir valores mediante un operador ternario",
    "Generar tarjetas con innerHTML y template literals",
    "Enviar un id en la URL y leerlo con URLSearchParams",
    "Buscar un objeto con find y completar el DOM",
    "Manejar errores mediante throw y catch",
]:
    bullet(item)
heading("Mapa del proceso")
code("""eventos.json
    -> app.js carga y recorre los datos
        -> index.html muestra una tarjeta por evento
            -> el enlace agrega ?id=identificador
                -> evento.js lee el id
                    -> busca el evento
                        -> evento.html muestra la ficha""")
heading("Estructura de los datos")
paragraph("El JSON contiene tres niveles de arrays. Cada nivel debe recorrerse antes de acceder al siguiente.")
code("""{
  "categorias": [
    {
      "nombre": "",
      "subcategorias": [
        {
          "nombre": "",
          "eventos": [
            {
              "id": "",
              "titulo": "",
              "artista": "",
              "fecha": "",
              "hora": "",
              "lugar": "",
              "precioBs": null,
              "descripcion": ""
            }
          ]
        }
      ]
    }
  ]
}""")
table(["Nivel", "Colección", "Contenido"], [
    ["1", "categorias", "Grupos generales"],
    ["2", "subcategorias", "Grupos específicos"],
    ["3", "eventos", "Objetos con los datos del evento"],
], [0.7, 1.7, 4.0])

page()
heading("Conceptos necesarios")
heading("Funciones asíncronas", 2)
paragraph("Una función asíncrona describe un proceso que debe esperar resultados sin bloquear el resto de la página. async permite usar await y hace que la función devuelva siempre una promesa.")
code("""async function cargarJson(ruta) {
    const respuesta = await fetch(ruta);

    if (!respuesta.ok) {
        throw new Error(§Error HTTP: ¤{respuesta.status}§);
    }

    return await respuesta.json();
}""")
table(["Forma", "Propósito", "Ejemplo"], [
    ["Función normal", "Devuelve un valor", "function sumar()"],
    ["Función async", "Puede esperar promesas", "async function cargar()"],
    ["Arrow function", "Callback de sintaxis breve", "error => error.message"],
], [1.35, 3.2, 1.95])
heading("Promesas y estados", 2)
paragraph("Una promesa representa un resultado que todavía no está disponible. Aquí aparecen tres: fetch produce una Response respuesta.json produce los datos y la función async representa el proceso completo.")
table(["Estado", "Significado"], [
    ["pending", "La operación todavía no termina"],
    ["fulfilled", "Terminó y produjo un resultado"],
    ["rejected", "Falló y produjo un error"],
], [1.4, 5.1])
heading("Estado HTTP y estado de la promesa", 2)
paragraph("fetch puede completarse aunque el servidor responda 404. Por eso se comprueba respuesta.ok y se genera un error con throw.")
code("""if (!respuesta.ok) {
    throw new Error(§No se puede cargar la agenda. Código ¤{respuesta.status}§);
}""")
heading("El bucle for of", 2)
paragraph("for of repite un bloque por cada valor de una colección. of no es una condición. Indica de qué colección se obtiene el valor de cada vuelta.")
code("""const colores = ["rojo", "verde", "azul"];

for (const color of colores) {
    console.log(color);
}""")
paragraph("Los tres for de app.js corresponden a categorías subcategorías y eventos. Se ejecutan después de convertir el JSON en un objeto JavaScript.")

page()
heading("Construcción de app js")
paragraph("app.js carga la agenda completa y crea una tarjeta por evento.")
steps_app = [
    ("Paso 1 seleccionar el contenedor", 'const listaEventos = document.querySelector(".lista_eventos");', "La variable representa el contenedor donde se insertarán tarjetas o un mensaje de error."),
    ("Paso 2 declarar la función de carga", 'async function cargarEventos() {\n    const respuesta = await fetch("./data/eventos.json");\n}', "fetch inicia la solicitud y await espera su promesa dentro de la función."),
    ("Paso 3 validar y convertir", 'if (!respuesta.ok) {\n    throw new Error(§No se puede cargar la agenda. Código ¤{respuesta.status}§);\n}\n\nconst agenda = await respuesta.json();', "respuesta.json transforma el texto JSON en arrays y objetos."),
    ("Paso 4 recorrer la jerarquía", 'for (const categoria of agenda.categorias) {\n    for (const subcategoria of categoria.subcategorias) {\n        for (const evento of subcategoria.eventos) {\n            // Construir una tarjeta\n        }\n    }\n}', "En el bucle interior están disponibles la categoría la subcategoría y el evento actual."),
    ("Paso 5 crear la tarjeta", 'const card = document.createElement("article");\ncard.className = "card-evento";\ncard.id = evento.id;', "createElement crea el nodo en memoria. La clase aplica el diseño y el id conserva el identificador."),
    ("Paso 6 calcular el precio", 'const entrada = evento.precioBs === 0\n    ? "Entrada libre"\n    : §Bs ¤{evento.precioBs}§;', "El ternario usa Entrada libre si la condición es verdadera y el precio si es falsa."),
]
for title, snippet, explanation in steps_app:
    heading(title, 2)
    code(snippet)
    paragraph(explanation)
code("""let entrada;
if (evento.precioBs === 0) {
    entrada = "Entrada libre";
} else {
    entrada = §Bs ¤{evento.precioBs}§;
}""")
paragraph("Este if else es equivalente al ternario.")

page()
heading("Paso 7 construir el contenido de la tarjeta", 2)
paragraph("innerHTML recibe un template literal. JavaScript reemplaza cada expresión entre signo de dólar y llaves por el valor del objeto actual.")
code("""card.innerHTML = §
    <p class="categoria-evento">
        ¤{categoria.nombre} - ¤{subcategoria.nombre}
    </p>
    <h3>¤{evento.titulo}</h3>
    <p>¤{evento.artista}</p>
§;""")
heading("Fecha y hora con significado semántico", 3)
code("""<p class="dato-evento">
    <i class="ph ph-calendar-dots" aria-hidden="true"></i>
    <time datetime="¤{evento.fecha}T¤{evento.hora}">
        ¤{evento.fecha} - ¤{evento.hora}
    </time>
</p>""")
for item in [
    "dato-evento permite al CSS alinear el icono y el texto",
    "ph-calendar-dots selecciona el icono de Phosphor",
    "aria-hidden oculta el icono decorativo a lectores de pantalla",
    "time identifica el contenido como fecha y hora",
    "datetime guarda un valor como 2026-12-04T20:30",
]:
    bullet(item)
heading("Enlace hacia el detalle", 3)
code("""<a href="./pages/evento.html?id=¤{encodeURIComponent(evento.id)}">
    Ver detalles del concierto
</a>""")
paragraph("El enlace abre evento.html y agrega el id como parámetro. encodeURIComponent prepara el valor para usarlo con seguridad en la URL.")
code("""evento.id = "laboratorio-sonoro"
URL generada = ./pages/evento.html?id=laboratorio-sonoro""")
heading("Paso 8 insertar la tarjeta", 2)
code("listaEventos.append(card);")
paragraph("append agrega la tarjeta al contenedor. Está dentro del tercer bucle y se ejecuta una vez por evento.")
heading("Paso 9 ejecutar y capturar errores", 2)
code("""cargarEventos().catch(error => {
    listaEventos.textContent = error.message;
});""")
paragraph("No se necesita then porque el camino exitoso ya está escrito con await dentro de cargarEventos. catch recibe el rechazo. error => es una arrow function usada como callback.")

page()
heading("Construcción de evento js")
paragraph("evento.js obtiene el id de la URL busca un único evento y completa evento.html.")
steps_detalle = [
    ("Paso 1 seleccionar los elementos", 'const estado = document.querySelector("#estado");\nconst detalle = document.querySelector("#detalle-evento");', "estado comunica carga o errores. detalle contiene la ficha oculta."),
    ("Paso 2 leer el id", 'const id = new URLSearchParams(window.location.search).get("id");', "location.search obtiene la sección que comienza con ?. URLSearchParams permite leer sus parámetros."),
    ("Paso 3 detenerse si falta el id", 'if (!id) {\n    estado.textContent = "Elige un concierto desde la agenda.";\n    return;\n}', "!id detecta null o una cadena vacía. return termina la función. Es una cláusula de guarda."),
    ("Paso 4 cargar la agenda", 'const respuesta = await fetch("../data/eventos.json");\nif (!respuesta.ok) {\n    throw new Error(§No se pudo cargar la agenda. Código ¤{respuesta.status}§);\n}\nconst agenda = await respuesta.json();', "La ruta usa ../ porque evento.html está dentro de pages."),
]
for title, snippet, explanation in steps_detalle:
    heading(title, 2)
    code(snippet)
    paragraph(explanation)
code("""URL        evento.html?id=ruido-de-altura
search     ?id=ruido-de-altura
get("id")  ruido-de-altura""")
heading("Paso 5 buscar el evento", 2)
code("""for (const categoria of agenda.categorias) {
    for (const subcategoria of categoria.subcategorias) {
        const evento = subcategoria.eventos.find(
            evento => evento.id === id
        );
    }
}""")
paragraph("Los bucles llegan a cada array de eventos. find devuelve el primer objeto cuyo id coincide. Si no encuentra ninguno devuelve undefined.")
code("""const eventoEncontrado = subcategoria.eventos.find(
    elemento => elemento.id === id
);""")
paragraph("Los nombres distintos aclaran que la variable externa guarda el resultado y elemento representa cada objeto probado.")

page()
heading("Paso 6 completar la ficha", 2)
paragraph("if evento solo se cumple cuando find devolvió un objeto. Sus propiedades se copian a la página.")
code("""if (evento) {
    document.title = §¤{evento.titulo} | Agenda de conciertos§;
    document.querySelector("#genero").textContent =
        §¤{categoria.nombre} · ¤{subcategoria.nombre}§;
    document.querySelector("#titulo").textContent = evento.titulo;
    document.querySelector("#artista").textContent = evento.artista;
    document.querySelector("#descripcion").textContent = evento.descripcion;
}""")
paragraph("textContent inserta texto sin interpretarlo como HTML.")
heading("Paso 7 configurar time", 2)
code("""const fechaHora = document.querySelector("#fecha-hora");
fechaHora.dateTime = §¤{evento.fecha}T¤{evento.hora}§;
fechaHora.textContent = §¤{evento.fecha} · ¤{evento.hora}§;""")
paragraph("dateTime define el valor estructurado y textContent la presentación visible.")
heading("Paso 8 mostrar y terminar", 2)
code("""estado.textContent = "";
detalle.hidden = false;
return;""")
paragraph("Se limpia el estado se revela la ficha y return evita seguir buscando.")
heading("Paso 9 informar que no existe", 2)
code('estado.textContent = "No se encontró el concierto solicitado.";')
paragraph("Esta línea está después de los bucles y solo se alcanza cuando ningún id coincide.")
heading("Paso 10 capturar errores", 2)
code("""cargarDetalle().catch(error => {
    estado.textContent = error.message;
});""")
paragraph("El patrón coincide con app.js. El éxito usa await dentro de la función y catch muestra cualquier rechazo.")

page()
heading("Comparación de los archivos")
table(["Aspecto", "app js", "evento js"], [
    ["Objetivo", "Mostrar todos los eventos", "Mostrar un evento"],
    ["Ruta JSON", "./data/eventos.json", "../data/eventos.json"],
    ["Recorrido", "Tres for of", "Dos for of y find"],
    ["Salida", "Crea article con innerHTML", "Completa nodos con textContent"],
    ["Identificador", "Lo agrega al enlace", "Lo obtiene de la URL"],
    ["Error", "Mensaje en listaEventos", "Mensaje en estado"],
], [1.25, 2.6, 2.65])
heading("Secuencia sugerida para la clase")
for item in [
    "Identificar los tres niveles de arrays en eventos.json",
    "Observar la promesa devuelta por fetch en la consola",
    "Construir app.js hasta obtener agenda con console.log",
    "Añadir los for of y registrar sus variables",
    "Crear una tarjeta mínima con título y artista",
    "Añadir precio fecha lugar iconos y enlace",
    "Comprobar el parámetro id en la URL",
    "Construir la guarda y la carga de evento.js",
    "Buscar con find y completar la ficha",
    "Probar caminos correctos y errores",
]:
    numbered(item)
heading("Preguntas de comprobación")
for item in [
    "¿Qué tres promesas aparecen durante la carga?",
    "¿Por qué un 404 requiere comprobar respuesta.ok?",
    "¿Qué representa categoria en cada vuelta?",
    "¿Qué devuelve find cuando no encuentra?",
    "¿Cómo decide el ternario el texto del precio?",
    "¿Qué diferencia existe entre datetime y el texto visible?",
    "¿Qué transporta el parámetro id?",
    "¿Por qué return evita trabajo innecesario?",
    "¿Por qué se usa catch sin then al final?",
]:
    bullet(item)
heading("Errores frecuentes")
table(["Problema", "Causa probable", "Comprobación"], [
    ["No aparecen tarjetas", "Ruta o fetch incorrectos", "Revisar Network y status"],
    ["El detalle no cambia", "Falta id en la URL", "Revisar location.search"],
    ["No se encuentra", "Los ids no coinciden", "Comparar evento.id e id"],
    ["Aparece undefined", "Falta una propiedad", "Validar los ocho campos"],
    ["JSON no carga", "Sintaxis inválida", "Validar comas y llaves"],
    ["fetch falla localmente", "No existe servidor HTTP", "Servir el proyecto"],
], [1.55, 2.5, 2.4])

page()
heading("Actividades para estudiantes")
for item in [
    "Cambiar Entrada libre por Gratis",
    "Mostrar una fecha legible sin cambiar datetime",
    "Agregar duración al JSON y a ambas páginas",
    "Reemplazar el ternario por if else",
    "Reescribir cargarEventos con then y comparar",
    "Probar evento.html sin id con id inexistente y con id válido",
]:
    bullet(item)
heading("Código completo de app js")
paragraph("El archivo completo sirve como referencia final. La actividad de clase debe reconstruirlo por etapas.")
code(APP)

page()
heading("Código completo de evento js")
paragraph("El archivo depende de los ids de eventos.json y de los elementos definidos en evento.html.")
code(EVENTO)
heading("Lista final de verificación")
for item in [
    "El JSON es válido y cada evento tiene id único",
    "Los scripts se cargan con defer",
    "Las rutas de fetch son correctas",
    "Los selectores coinciden con el HTML",
    "Cada tarjeta envía el id codificado",
    "La ficha muestra el evento y comunica errores",
    "La consola no muestra errores JavaScript",
]:
    bullet(item)

doc.core_properties.title = "Guía para construir app js y evento js"
doc.core_properties.subject = "Carga de JSON generación de tarjetas y detalle de eventos"
doc.core_properties.author = ""
doc.core_properties.keywords = "JavaScript JSON async await fetch promesas for of URLSearchParams"
doc.save(OUT)
print(OUT)

