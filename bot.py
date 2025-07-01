from fastapi import FastAPI, File, UploadFile, Form
import uvicorn
from telegram import Update
from telegram import Application, CommandHandler, MessageHandler, filters
import os
import httpx
from config import TELEGRAM_TOKEN, VIDEO_PROCESSING_SERVICE_URL, PRESENTATION_GENERATION_SERVICE_URL, BOT_REQUEST_URL

app = FastAPI()

@app.post("/handle_video/")
async def handle_video(file: UploadFile = File(...)):
    file_location = f"temp_{file.filename}"
    with open(file_location, "wb+") as file_object:
        file_object.write(file.file.read())

    try:
        async with httpx.AsyncClient(follow_redirects=True) as client:
            print(f"Uploading video to {VIDEO_PROCESSING_SERVICE_URL}")
            with open(file_location, 'rb') as video_file:
                response = await client.post(VIDEO_PROCESSING_SERVICE_URL, files={'file': video_file})
                print(f"Response from video processing service: {response.text}")
                data = response.json()

            print(f"Sending data to {PRESENTATION_GENERATION_SERVICE_URL}")
            response = await client.post(PRESENTATION_GENERATION_SERVICE_URL, json=data)
            response.raise_for_status()
            print(f"Response from presentation generation service: {response.text}")
            presentation_file_url = response.json().get('presentation_url')

        return {"presentation_url": presentation_file_url}
    except httpx.HTTPStatusError as e:
        print(f"HTTP error occurred: {e.response.status_code}, {e.response.text}")
        return {"error": f"HTTP error occurred: {e.response.status_code}, {e.response.text}"}
    except httpx.RequestError as e:
        print(f"Request error occurred: {str(e)}")
        return {"error": f"Request error occurred: {str(e)}"}
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return {"error": f"An error occurred: {str(e)}"}
    finally:
        os.remove(file_location)

async def start(update: Update, context):
    await update.message.reply_text('Приветствую! Отправляйте видео ролик!')

async def help(update: Update, context):
    await update.message.reply_text('Помощь!')

async def video_handler_bot(update: Update, context):
    file = await update.message.video.get_file()
    file_path = await file.download_to_drive()

    try:
        async with httpx.AsyncClient(follow_redirects=True) as client:
            print(f"Uploading video to bot request URL: {BOT_REQUEST_URL}")
            with open(file_path, "rb") as video_file:
                response = await client.post(BOT_REQUEST_URL, files={"file": video_file})
                response.raise_for_status()
                response_json = response.json()
                print(f"Response from bot request URL: {response.text}")

        if "presentation_url" in response_json:
            presentation_url = response_json["presentation_url"]
            
            # Отправка файла презентации пользователю
            await context.bot.send_document(
                chat_id=update.effective_chat.id,
                document=open(presentation_url, 'rb')
            )
            
            # Очистка загруженного файла презентации
            os.remove(presentation_url)
        else:
            error_message = response_json.get('error', 'Unknown error occurred.')
            await update.message.reply_text(f"Error: {error_message}")
    except httpx.HTTPStatusError as e:
        await update.message.reply_text(f"HTTP error occurred: {e.response.status_code}, {e.response.text}")
    except httpx.RequestError as e:
        await update.message.reply_text(f"Request error occurred: {str(e)}")
    except Exception as e:
        await update.message.reply_text(f"An error occurred: {str(e)}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)


def run_fastapi():
    uvicorn.run(app, host="127.0.0.1", port=8000)

def main():
    application = Application.builder().token(TELEGRAM_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help))
    application.add_handler(MessageHandler(filters.VIDEO, video_handler_bot))

    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    import threading

    # Запуск FastAPI в отдельном потоке
    fastapi_thread = threading.Thread(target=run_fastapi)
    fastapi_thread.start()

    # Запуск Telegram бота
    main()
