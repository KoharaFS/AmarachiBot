import threading
from queue import Queue
import gc
import speech_recognition as sr
import edge_tts
import asyncio
import tempfile
from playsound import playsound
from pydub import AudioSegment
from pydub.playback import play
import tempfile
import nltk
import string
import google.generativeai as genai
import os
import openpyxl
import time
from dotenv import load_dotenv
from datetime import datetime
from nltk.corpus import stopwords
from threading import Thread

# Load environment variables for Gemini API
load_dotenv()
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# Initialize Gemini Model
model = genai.GenerativeModel("gemini-2.0-flash")

# Download NLTK Data if not present
try:
    nltk.data.find('corpora/stopwords.zip')
    nltk.data.find('tokenizers/punkt.zip')
except LookupError:
    nltk.download('stopwords')
    nltk.download('punkt')

# Initialize Recognizer
recognizer = sr.Recognizer()

# Initialize Queue for TTS
tts_queue = Queue()

# Load or create Excel file
try:
    workbook = openpyxl.load_workbook('SpeechData.xlsx')
    sheet = workbook.active
except FileNotFoundError:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Speech Timestamp", "Correction Timestamp", "Recognized Speech", "Completed Sentence", "Total Processing Time (s)", "API Processing Time (s)", "Speech Processing Time (s)"])

# Speech-to-Text Function
def speech_to_text():
    with sr.Microphone() as source:
        print("I'm listening...")
        recognizer.adjust_for_ambient_noise(source)
        try:
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=10)
            text = recognizer.recognize_google(audio).lower()
            print("You said:", text)
            return text
        except sr.UnknownValueError:
            print("Sorry, could not understand the audio.")
        except sr.RequestError:
            print("Could not request results, check your internet connection.")
        except sr.WaitTimeoutError:
            print("No speech detected, stopping...")
    return ""

# Text Preprocessing
def preprocess_text(text):
    text = text.lower()
    text = text.translate(str.maketrans('', '', string.punctuation))
    words = nltk.word_tokenize(text)
    try:
        stop_words = set(stopwords.words('english'))
        filtered_words = [word for word in words if word not in stop_words]
    except LookupError:
        return text
    return ' '.join(filtered_words)

# Gemini API Autocomplete Function
def autofill_sentence_with_retry(broken_sentence, retries=3):
    if not broken_sentence:
        return "No input detected.", 0.0
    attempt = 0
    while attempt < retries:
        try:
            prompt = (
                "You are a speech support assistant designed to help aphasia patients complete broken sentences. "
                "You must only return **one full corrected version** of the sentence. "
                "Do not explain anything, do not ask questions, and do not add extra lines. "
                "Do not end the task early. Do not say 'okay' or 'I understand'. Just fix the sentence completely.\n\n"
                f"Incomplete sentence: {broken_sentence}\n\n"
                "Return only the corrected version:"
            )
            gemini_start_time = time.perf_counter()
            response = model.generate_content(prompt)
            gemini_end_time = time.perf_counter()
            gemini_processing_time = round(gemini_end_time - gemini_start_time, 6)
            completed_sentence = response.candidates[0].content.parts[0].text.strip()
            return completed_sentence if completed_sentence else broken_sentence, gemini_processing_time
        except Exception as e:
            print(f"Error with Gemini API (attempt {attempt+1}): {e}")
            attempt += 1
            time.sleep(2)
    print("Max retries reached, returning original sentence.")
    return broken_sentence, 0.0

async def edge_tts_speak(text):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmpfile:
            temp_path = tmpfile.name

        communicate = edge_tts.Communicate(text=text, voice="en-US-AriaNeural")
        await communicate.save(temp_path)

        audio = AudioSegment.from_file(temp_path, format="mp3")
        play(audio)

        os.remove(temp_path)
    except Exception as e:
        print(f"Edge TTS error: {e}")

def tts_worker():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    while True:
        text = tts_queue.get()
        if text is None:
            break
        try:
            loop.run_until_complete(edge_tts_speak(text))
        except Exception as e:
            print(f"TTS Worker Error: {e}")
        finally:
            tts_queue.task_done()
            
# Main Function
def main():
    print("Loading...")

    # Start TTS worker thread
    tts_thread = Thread(target=tts_worker, daemon=True)
    tts_thread.start()

    update_made = False

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)

    try:
        while True:
            print("[DEBUG] Starting speech recognition...")
            speech_start_time = time.perf_counter()
            spoken_text = speech_to_text()
            speech_end_time = time.perf_counter()
            speech_processing_time = round(speech_end_time - speech_start_time, 6)

            if not spoken_text:
                continue

            if spoken_text == "zero":
                print("Exiting program...")
                tts_queue.put(None)  # Signal TTS worker to stop
                break

            speech_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            processed_text = preprocess_text(spoken_text)
            corrected_sentence, gemini_processing_time = autofill_sentence_with_retry(processed_text)
            correction_timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            total_processing_time = round(speech_processing_time + gemini_processing_time, 6)

            print("Corrected Sentence:", corrected_sentence)
            print(f"Gemini Processing Time: {gemini_processing_time} seconds")
            print(f"Speech Processing Time: {speech_processing_time} seconds")
            print(f"Total Processing Time: {total_processing_time} seconds")

            # Queue the corrected sentence for TTS
            tts_queue.put(corrected_sentence)

            # Log to Excel
            sheet.append([
                speech_timestamp,
                correction_timestamp,
                spoken_text,
                corrected_sentence,
                total_processing_time,
                gemini_processing_time,
                speech_processing_time
            ])
            update_made = True

            # Garbage collection and cooldown
            time.sleep(1)
            gc.collect()

    except KeyboardInterrupt:
        print("Program interrupted. Exiting...")
        tts_queue.put(None)  # Ensure TTS worker stops

    finally:
        if update_made:
            workbook.save('SpeechData.xlsx')
            print("Data saved to SpeechData.xlsx.")

if __name__ == "__main__":
    main()
