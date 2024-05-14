import speech_recognition as sr
import time


languages = ["en-EN", "de-DE"]
input_collection = list()

rec = sr.Recognizer()
mic = sr.Microphone()


def capture_input(prompt="Next"):
    with mic as source:
        # rec.adjust_for_ambient_noise(source)
        print(prompt)
        audio = rec.listen(source)
        time.sleep(3)
    return audio


def get_command(prompt="next, finished or print"):
    print("Enter command")
    command = create_text_from_audio(capture_input(prompt))
    return command.lower()


def create_text_from_audio(audio):
    try:
        text = rec.recognize_google(audio)
        print("Google Speech Recognition thinks you said " + text)
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        text = "Unrecognized"
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
        text = "Error"
    return text


def create_entry():
    entry = list()
    prompts = ["Date", "Category", "Amount"]
    print("Creating new entry: ")
    for i in range(3):
        time.sleep(3)
        audio = capture_input(prompts[i])
        text = create_text_from_audio(audio)
        entry.append(text)
    return entry


command = get_command()
while command not in ["next", "finished", "print"]:
    command = get_command("Come again?")
while command == "next":
    new_entry = create_entry()
    print("Adding to collection:", new_entry)
    input_collection.append(new_entry)
    command = get_command()
if command == "finished":
    print(input_collection)
