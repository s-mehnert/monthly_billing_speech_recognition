# Speech recognition for Monthly Billing

In order to speed up the process of the monthly exercise of entering all dates, amounts and according categories of receipts in a spending spreadsheet, this project aims at converting data from a previously recorded audio file to text and then format and store the relevant data in an excel spreadsheet for further usage.

## Installation

For this project the following Python libraries need to be installed:
- SpeechRecognition
- PyAudio
- Openpyxl

The full list of dependencies can be found in and installed directly from the requirements document.

```bash
pip install -r requirements.txt
```

## Usage

Make a recording and save the file with the .wav extension in the same folder as the Python script. Note, that you will need to change the filename in the python script accordingly. Currently, the script is pointing to "test3.wav" which is a short sample recording of mine to test and demonstrate the workings of the script. Note, that in order for the conversion to work properly you need to be speaking English and use the following structure:
1. Name of the month
2. "next" --> Date, category and amount of receipt
3. repeat for all the receipts you have for the month
***make sure to start each receipt with the keyword "next"***

Then run the script.

```bash
python monthly_billing_speech_recognition__python.py
```
When the script is finished open billing.xlsx. You should see a new sheet for the month you recorded containing the data you recorded in three columns as well as the calculation of the total sum.
Note that depending on the quality of the audio, not all data will be placed into the excel file correctly. However, entries (=receipts) the script couldn't correctly split up will nevertheless show up as strings in the excel file to make the necessary manual corrections easier. Numbers mostly show up correctly in the Amount column, however dates do not.

## Planned for version 2:
- Date conversion
- Checking for unreasonably high amounts (can occur when the decimal point is not recognized)
- Improving style formatting to excel sheet from a template

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.


## License

[MIT](https://choosealicense.com/licenses/mit/)