import speech_recognition as sr
import win32com.client
import time
import re


# ============================================================
# WORD TO NUMBER CONVERSION (ENGLISH + SINHALA)
# ============================================================

number_words = {
    # English
    "zero": 0,
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
    "twenty": 20,
    "thirty": 30,
    "forty": 40,
    "fifty": 50,
    "sixty": 60,
    "seventy": 70,
    "eighty": 80,
    "ninety": 90,

    # Sinhala
    "බිංදුව": 0, "බින්දුව": 0, "බිංදු": 0,
    "එක": 1, "එකක්": 1,
    "දෙක": 2, "දෙකක්": 2,
    "තුන": 3, "තුනක්": 3,
    "හතර": 4, "හතරක්": 4,
    "පහ": 5, "පහක්": 5, "පන්": 5,
    "හය": 6, "හයක්": 6, "හයසිය": 6,
    "හත": 7, "හතක්": 7, "හත්": 7,
    "අට": 8, "අටක්": 8, "අටසිය": 8,
    "නවය": 9, "නමය": 9, "නවයක්": 9, "නමයක්": 9, "නවසිය": 9, "නමසිය": 9,
    "දහය": 10, "දහයක්": 10,
    "එකොළොස": 11, "එකොළොස්": 11, "එකොළහ": 11,
    "දොළොස": 12, "දොළොස්": 12, "දොළහ": 12,
    "තෙළෙස": 13, "තෙළෙස්": 13, "දහතුන": 13,
    "දහහතර": 14,
    "පහළොස": 15, "පහළොස්": 15, "පහළොව": 15, "පසළොස්": 15,
    "දහසය": 16,
    "දහහත": 17,
    "දහඅට": 18,
    "දහනවය": 19, "දහනමය": 19,
    "විස්ස": 20, "විසි": 20,
    "තිහ": 30, "තිස්": 30,
    "හතළිහ": 40, "හතළිස්": 40,
    "පනහ": 50, "පනස්": 50,
    "හැට": 60, "හැටක්": 60,
    "හැත්තෑව": 70, "හැත්තෑ": 70,
    "අසූව": 80, "අසූ": 80,
    "අනූව": 90, "අනූ": 90,

    # Compound Sinhala hundred prefix forms
    "එක්සිය": 100, "දෙසිය": 200, "තුන්සිය": 300, "හාරසිය": 400, "පන්සිය": 500, "පන්සියය": 500
}

scale_words = {
    # English
    "hundred": 100,
    "thousand": 1000,
    "million": 1000000,
    "billion": 1000000000,

    # Sinhala
    "සියය": 100, "සිය": 100, "සියයක්": 100, "සියයම": 100,
    "දහස": 1000, "දහසක්": 1000, "දාහ": 1000, "දාහක්": 1000, "දහස්": 1000,
    "ලක්ෂය": 100000, "ලක්ෂ": 100000, "ලක්ෂයක්": 100000,
    "මිලියන": 1000000, "මිලියනය": 1000000, "මිලියනයක්": 1000000,
    "බිලියන": 1000000000, "බිලියනය": 1000000000, "බිලියනයක්": 1000000000
}

SINHALA_DIGITS = {
    '෦': '0', '෧': '1', '෨': '2', '෩': '3', '෪': '4',
    '෫': '5', '෬': '6', '෭': '7', '෮': '8', '෯': '9'
}


def normalize_text(text):
    """Normalize text including Sinhala digits to standard ASCII digits."""
    for s_digit, ascii_digit in SINHALA_DIGITS.items():
        text = text.replace(s_digit, ascii_digit)
    return text


def words_to_number(text):
    """
    Convert spoken English or Sinhala number words into a number.

    Examples:
        one hundred twenty five -> 125
        තුන -> 3
        සියය -> 100
    """

    text = normalize_text(text.lower())
    text = text.replace("-", " ")
    text = text.replace(",", " ")

    words = text.split()

    # Remove "and" and Sinhala connector "සහ"
    words = [word for word in words if word not in ["and", "සහ"]]

    # If the speech already contains a numeric value
    joined = "".join(words)

    if re.fullmatch(r"\d+(\.\d+)?", joined):
        return float(joined) if "." in joined else int(joined)

    total = 0
    current = 0

    for word in words:

        if word in number_words:
            current += number_words[word]

        elif word in ["hundred", "සියය", "සිය", "සියයක්"]:
            if current == 0:
                current = 1
            current *= 100

        elif word in scale_words:
            if current == 0:
                current = 1

            total += current * scale_words[word]
            current = 0

        elif word in ["point", "decimal", "දශම"]:
            # Decimal numbers handled separately
            break

    result = total + current

    # Handle decimal part
    decimal_markers = [m for m in ["point", "decimal", "දශම"] if m in words]
    if decimal_markers:
        position = words.index(decimal_markers[0])

        integer_words = words[:position]
        decimal_words = words[position + 1:]

        integer_value = words_to_number(" ".join(integer_words)) if integer_words else 0

        decimal_digits = ""

        for word in decimal_words:

            if word in number_words:
                value = number_words[word]

                if value < 10:
                    decimal_digits += str(value)
                else:
                    decimal_digits += str(value)

            elif word.isdigit():
                decimal_digits += word

        if decimal_digits:
            return float(str(integer_value) + "." + decimal_digits)

    return result


# ============================================================
# EXTRA NUMBER PROCESSING
# ============================================================

def convert_number(text):

    text = normalize_text(text.lower().strip())

    # Remove common filler words
    fillers = ["please", "enter", "කරුණාකර", "ඇතුළත් කරන්න", "ඇතුලත් කරන්න"]
    for filler in fillers:
        text = text.replace(filler, "")
    text = text.strip()

    # Direct numeric speech
    match = re.search(r"\d+(?:\.\d+)?", text)

    if match:
        value = match.group()

        if "." in value:
            return float(value)
        else:
            return int(value)

    # Word-based number
    try:
        value = words_to_number(text)

        if value is not None:
            return value
    except Exception:
        pass

    return None


def recognize_speech(recognizer, audio):
    """
    Recognize speech trying English first, then Sinhala.
    Returns recognized text or None if failed.
    """
    # Try English (en-US)
    try:
        text = recognizer.recognize_google(audio, language="en-US")
        if text:
            return text.lower().strip()
    except Exception:
        pass

    # Try Sinhala (si-LK)
    try:
        text = recognizer.recognize_google(audio, language="si-LK")
        if text:
            return text.lower().strip()
    except Exception:
        pass

    return None


# ============================================================
# MAIN PROGRAM
# ============================================================

if __name__ == "__main__":
    # ---------- Connect to Excel ----------
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("ERROR: Excel is not open.")
        print("Please open Excel and select the first cell.")
        input("Press Enter to exit...")
        exit()

    # ---------- Get the active worksheet and cell ----------
    sheet = excel.ActiveSheet
    cell = excel.ActiveCell

    print("=" * 60)
    print("       VOICE NUMBER ENTRY FOR EXCEL (ENGLISH & SINHALA)")
    print("=" * 60)

    print(f"Excel Sheet : {sheet.Name}")
    print(f"Starting Cell : {cell.Address}")

    print()
    print("Commands / විධානයන්:")
    print("  Say a number / අංකයක් කියන්න -> Enter the number")
    print("  'next' / 'ඊළඟ'                -> Move to next row (downwards)")
    print("  'back' / 'ආපසු' me             -> Move to previous row (upwards)")
    print("  'clear' / 'මකන්න'             -> Clear current cell")
    print("  'stop' / 'නවත්වන්න'          -> Stop the program")
    print()
    print("Example:")
    print("  Say: one hundred twenty five OR සියය විසි පහ")
    print("  Excel: 125")
    print("=" * 60)

    # ---------- Speech recognizer ----------
    recognizer = sr.Recognizer()

    # Adjust microphone sensitivity
    recognizer.energy_threshold = 300
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.8

    # ---------- Microphone ----------
    try:
        microphone = sr.Microphone()
    except Exception:
        print("ERROR: Microphone could not be found.")
        print("Check Windows microphone settings.")
        input("Press Enter to exit...")
        exit()

    current_cell = cell

    # Adjust microphone for ambient noise before starting the main loop
    print()
    print("Adjusting microphone for background noise...")
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)

    print("Ready!")
    print()
    print(f"Current Excel cell: {current_cell.Address}")
    print("Start speaking (English or Sinhala)...")
    print()

    while True:

        try:

            with microphone as source:

                print("Listening...")

                audio = recognizer.listen(
                    source,
                    timeout=10,
                    phrase_time_limit=8
                )

            print("Processing...")

            text = recognize_speech(recognizer, audio)

            if not text:
                print("Could not understand. Please try again.")
                print()
                continue

            text = text.lower().strip()

            print(f"You said: {text}")

            # ------------------------------------------------
            # STOP
            # ------------------------------------------------

            if text in [
                "stop", "exit", "quit", "close", "finish",
                "නවත්වන්න", "නවත්තන්න", "නවතන්න", "අයින් වෙන්න", "නතර කරන්න"
            ]:

                print()
                print("Voice entry stopped.")
                break

            # ------------------------------------------------
            # NEXT
            # ------------------------------------------------

            if text in [
                "next", "next row", "next cell",
                "ඊළඟ", "ඊලඟ", "ඊළඟ පේළිය", "ඊළඟ කොටුව", "පහළට", "පහලට"
            ]:

                # Move vertically downwards (increase Row number by 1 in the same Column)
                current_cell = sheet.Cells(current_cell.Row + 1, current_cell.Column)
                current_cell.Select()

                print(f"Moved to {current_cell.Address}")
                print()

                continue

            # ------------------------------------------------
            # BACK
            # ------------------------------------------------

            if text in [
                "back", "previous", "previous row", "previous cell",
                "ආපසු", "කලින්", "කලින් පේළිය", "ඉහළට", "උඩට"
            ]:

                # Move vertically upwards (decrease Row number by 1, minimum row 1)
                prev_row = max(1, current_cell.Row - 1)
                current_cell = sheet.Cells(prev_row, current_cell.Column)
                current_cell.Select()

                print(f"Moved to {current_cell.Address}")
                print()

                continue

            # ------------------------------------------------
            # CLEAR
            # ------------------------------------------------

            if text in [
                "clear", "delete", "clear cell",
                "මකන්න", "අයින් කරන්න", "ක්ලියර්"
            ]:

                current_cell.ClearContents()

                print(f"Cleared {current_cell.Address}")
                print()

                continue

            # ------------------------------------------------
            # CONVERT NUMBER
            # ------------------------------------------------

            number = convert_number(text)

            if number is not None:

                # Enter number into Excel
                current_cell.Value = number

                print(
                    f"Entered {number} into "
                    f"{current_cell.Address}"
                )

                # Automatically move to next row vertically (downwards)
                # sheet.Cells(Row, Column): increment Row by 1 to move down to the next row
                current_cell = sheet.Cells(current_cell.Row + 1, current_cell.Column)

                current_cell.Select()

                print(
                    f"Next cell: "
                    f"{current_cell.Address}"
                )

                print()

            else:

                print(
                    "I could not identify a number "
                    "from your speech."
                )

                print("Please try again.")
                print()

        except sr.WaitTimeoutError:

            print("No speech detected.")
            print()

        except KeyboardInterrupt:

            print()
            print("Program stopped.")
            break

        except Exception as error:

            print()
            print("ERROR:", error)
            print()

            time.sleep(1)


    print()
    print("=" * 60)
    print("Program finished.")
    print("=" * 60)
    input("Press Enter to close...")
