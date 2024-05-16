import speech_recognition as sr
import time
import openpyxl


languages = ["en-EN", "de-DE"]
input_collection = list()

current_year = "2024"

rec = sr.Recognizer()
mic = sr.Microphone()

wb = openpyxl.load_workbook("billing_2.xlsx")



def create_text_from_audio(audio_file):
    try:
        billing_audio = sr.AudioFile(audio_file)
        with billing_audio as source:
            rec.adjust_for_ambient_noise(source)
            audio = rec.record(source)
        text = rec.recognize_google(audio)
        print("Google Speech Recognition thinks you said " + text)
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        text = "Unrecognized"
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
        text = "Error"
    return text


def sort_entries(entry_list):
    check_again = list()
    cleared = list()
    for entry in entry_list:
        if len(entry.split()) != 3:
            check_again.append(entry)
            print("Invalid entry:", entry)
        else:
            valid_entry = entry.split()
            cleared.append(valid_entry)
            print("Valid entry:", valid_entry)
    return cleared, check_again


def convert_month(m):  # add try except block
    months = {"january": "01", "february": "02",  "march": "03", "april": "04", "may": "05",
              "june": "06", "july": "07", "august": "08", "september": "09", "october": "10",
              "november": "11", "december": "12"}
    return months[m.lower()]


def create_full_date(d, m, y):
    return y + convert_month(m) + d


billing_data = create_text_from_audio("test3.wav")
bd_listed = billing_data.split("next")
print(billing_data)
print(bd_listed)

valid_entries, invalid_entries = sort_entries(bd_listed)
print("Valid entries:", valid_entries)
print("Check again:", invalid_entries)

current_month = invalid_entries.pop(0).strip()
print(current_month)
ws = wb.create_sheet(current_month)

ws["A1"] = "Date"
ws["B1"] = "Category"
ws["C1"] = "Amount"

for entry in valid_entries:
    row_number = ws.max_row + 1
    ws["A" + str(row_number)] = create_full_date(entry[0].strip(), current_month, current_year)
    ws["B" + str(row_number)] = entry[1].strip()
    ws["C" + str(row_number)] = entry[2].strip()

for entry in invalid_entries:
    row_number = ws.max_row + 1
    ws["E" + str(row_number)] = entry

wb.save("billing_2.xlsx")
print("Data saved to file.")
wb.close()
    





