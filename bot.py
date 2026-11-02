import logging
import sqlite3
import os
from datetime import datetime
from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, CommandHandler, MessageHandler, filters
import openpyxl  # Librería para generar archivos Excel

# Configurar base de datos SQLite local
def init_db():
    conn = sqlite3.connect('registro_givit.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            horas REAL,
            pedidos INTEGER
        )
    ''')
    conn.commit()
    conn.close()

init_db()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "¡Hola, Enmanuel! 🛵📊 Aquí tienes tu centro de control para Givit.\n\n"
        "📝 **Cómo registrar:**\n"
        "Solo escríbeme los números así: `[horas] [pedidos]` (Ej: `6.5 14`).\n\n"
        "📊 **Comandos disponibles:**\n"
        "• /resumen - Ver totales generales.\n"
        "• /excel - Te genera y envía un archivo Excel con todos tus datos.\n"
        "• /borrar_ultimo - Elimina el último registro por error.\n\n"
        "💬 También puedes preguntarme cosas como: *'¿Cuánto hice en octubre?'* o *'¿Qué hice el 2026-10-15?'*."
    )

async def guardar_registro(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip()
    
    # Si el mensaje parece una pregunta en texto normal, pasarlo al sistema de consultas en vez de error
    if any(palabra in texto.lower() for palabra in ["cuanto", "cuánto", "hice", "mes", "octubre", "noviembre", "diciembre", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre"]):
        await responder_consulta(update, context)
        return

    partes = texto.split()
    if len(partes) != 2:
        await update.message.reply_text("⚠️ Formato incorrecto. Envíalo así: `[horas] [pedidos]` (Ej: `7.5 12`) o usa los comandos.")
        return

    try:
        horas = float(partes[0].replace(',', '.'))
        pedidos = int(partes[1])
    except ValueError:
        await update.message.reply_text("⚠️ Asegúrate de usar números válidos. Ejemplo: `6.5 15`")
        return

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")

    conn = sqlite3.connect('registro_givit.db')
    cursor = conn.cursor()
    cursor.execute("INSERT INTO registros (fecha, horas, pedidos) VALUES (?, ?, ?)", (fecha_hoy, horas, pedidos))
    conn.commit()
    conn.close()

    await update.message.reply_text(
        f"✅ **¡Guardado con éxito!** ({fecha_hoy})\n"
        f"⏱️ Horas: {horas}h\n"
        f"📦 Pedidos: {pedidos}\n\n"
        f"Usa /resumen para ver tus totales o /excel para descargar el archivo."
    )

async def ver_resumen(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect('registro_givit.db')
    cursor = conn.cursor()
    cursor.execute("SELECT SUM(horas), SUM(pedidos), COUNT(*) FROM registros")
    resultado = cursor.fetchone()
    conn.close()

    total_horas = resultado[0] if resultado[0] is not None else 0.0
    total_pedidos = resultado[1] if resultado[1] is not None else 0
    dias_trabajados = resultado[2] if resultado[2] is not None else 0

    await update.message.reply_text(
        f"📊 **Resumen Total Acumulado**\n\n"
        f"📅 Días registrados: {dias_trabajados}\n"
        f"⏱️ Total horas: **{round(total_horas, 2)} h**\n"
        f"📦 Total pedidos: **{total_pedidos}**"
    )

async def generar_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect('registro_givit.db')
    cursor = conn.cursor()
    cursor.execute("SELECT fecha, horas, pedidos FROM registros ORDER BY fecha ASC")
    filas = cursor.fetchall()
    conn.close()

    if not filas:
        await update.message.reply_text("⚠️ No hay registros guardados todavía para exportar.")
        return

    # Crear archivo Excel usando openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Control Givit"

    # Cabeceras
    ws.append(["Fecha", "Horas Trabajadas", "Pedidos Realizados"])

    # Añadir datos
    for fila in filas:
        ws.append(list(fila))

    # Guardar temporalmente
    nombre_archivo = "Control_Horas_Givit.xlsx"
    wb.save(nombre_archivo)

    # Enviar archivo por Telegram
    with open(nombre_archivo, 'rb') as f:
        await update.message.reply_document(
            document=f,
            caption="📁 Aquí tienes tu archivo Excel actualizado con todos los registros de Givit."
        )

    # Limpiar archivo local opcionalmente
    if os.path.exists(nombre_archivo):
        os.remove(nombre_archivo)

async def responder_consulta(update: Update, context: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.lower()
    conn = sqlite3.connect('registro_givit.db')
    cursor = conn.cursor()

    # Mapeo simple de meses en texto a número de mes
    meses = {
        "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
        "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
        "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12"
    }

    mes_encontrado = None
    for nombre_mes, num_mes in meses.items():
        if nombre_mes in texto:
            mes_encontrado = num_mes
            break

    if mes_encontrado:
        # Filtrar por mes en el año actual (2026)
        cursor.execute("SELECT SUM(horas), SUM(pedidos), COUNT(*) FROM registros WHERE fecha LIKE ?", (f"2026-{mes_encontrado}%",))
        res = cursor.fetchone()
        conn.close()

        th = res[0] if res[0] is not None else 0.0
        tp = res[1] if res[1] is not None else 0
        dias = res[2] if res[2] is not None else 0

        await update.message.reply_text(
            f"📅 **Datos para el mes seleccionado:**\n\n"
            f"⏱️ Total horas: {round(th, 2)} h\n"
            f"📦 Total pedidos: {tp}\n"
            f"🗓️ Días trabajados: {dias}"
        )
    else:
        conn.close()
        await update.message.reply_text("🤔 No he entendido bien tu pregunta. Puedes usar /resumen para ver el total, o pedirme el Excel con /excel.")

async def borrar_ultimo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    conn = sqlite3.connect('registro_givit.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, fecha, horas, pedidos FROM registros ORDER BY id DESC LIMIT 1")
    ultimo = cursor.fetchone()
    
    if ultimo:
        cursor.execute("DELETE FROM registros WHERE id = ?", (ultimo[0],))
        conn.commit()
        conn.close()
        await update.message.reply_text(f"🗑️ Borrado el último registro del {ultimo[1]} ({ultimo[2]}h, {ultimo[3]} pedidos).")
    else:
        conn.close()
        await update.message.reply_text("⚠️ No hay registros guardados todavía.")

def main():
    TOKEN = "9055369271:AAHIM5zIGqhFw121pqjLvpFDKi7LAmW809A"
    
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("resumen", ver_resumen))
    app.add_handler(CommandHandler("excel", generar_excel))
    app.add_handler(CommandHandler("borrar_ultimo", borrar_ultimo))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), guardar_registro))

    print("🤖 Bot avanzado de Givit en marcha...")
    app.run_polling()

if __name__ == '__main__':
    main()