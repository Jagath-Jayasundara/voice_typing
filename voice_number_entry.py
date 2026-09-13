import speech_recognition as sr
import win32com.client
import time
import re


# ============================================================
# WORD TO NUMBER CONVERSION
# ============================================================

number_words = {
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
    "ninety": 90
}

scale_words = {
    "hundred": 100,
    "thousand": 1000,
    "million": 1000000,
    "billion": 1000000000
}


def words_to_number(text):
    """
    Convert spoken English number words into a number.

    Examples:
        one hundred twenty five -> 125
        two thousand five hundred -> 2500
        twenty five thousand -> 25000
    """

    text = text.lower()
    text = text.replace("-", " ")
    text = text.replace(",", " ")

    words = text.split()

    # Remove "and"
    words = [word for word in words if word != "and"]

    # If the speech already contains a numeric value
    joined = "".join(words)

    if re.fullmatch(r"\d+(\.\d+)?", joined):
        return float(joined) if "." in joined else int(joined)

    total = 0
    current = 0

    for word in words:

        if word in number_words:
            current += number_words[word]

        elif word == "hundred":
            if current == 0:
                current = 1
            current *= 100

        elif word in ["thousand", "million", "billion"]:
            if current == 0:
                current = 1

            total += current * scale_words[word]
            current = 0

        elif word in ["point", "decimal"]:
            # Decimal numbers handled separately
            break

    result = total + current

    # Handle decimal part
    if "point" in words or "decimal" in words:

        if "point" in words:
            position = words.index("point")
        else:
            position = words.index("decimal")

        integer_words = words[:position]
        decimal_words = words[position + 1:]

        integer_value = words_to_number(" ".join(integer_words))

        decimal_digits = ""

        for word in decimal_words:

            if word in number_words:
                value = number_words[word]

                if value < 10:
                    decimal_digits += str(value)
                else:
                    # For example "twenty five"
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

    text = text.lower().strip()

    # Remove common filler words
    text = text.replace("please", "")
    text = text.replace("enter", "")
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
    print("       VOICE NUMBER ENTRY FOR EXCEL")
    print("=" * 60)

    print(f"Excel Sheet : {sheet.Name}")
    print(f"Starting Cell : {cell.Address}")

    print()
    print("Commands:")
    print("  Say a number      -> Enter the number")
    print("  'next'            -> Move to next row (downwards)")
    print("  'back'            -> Move to previous row (upwards)")
    print("  'clear'           -> Clear current cell")
    print("  'stop'            -> Stop the program")
    print()
    print("Example:")
    print("  Say: one hundred twenty five")
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
    print("Start speaking...")
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

            try:

                text = recognizer.recognize_google(
                    audio,
                    language="en-US"
                )

            except sr.UnknownValueError:

                print("Could not understand. Please try again.")
                print()
                continue

            except sr.RequestError:

                print("Speech recognition service is unavailable.")
                print("Please check your Internet connection.")
                print()
                continue

            text = text.lower().strip()

            print(f"You said: {text}")

            # ------------------------------------------------
            # STOP
            # ------------------------------------------------

            if text in [
                "stop",
                "exit",
                "quit",
                "close",
                "finish"
            ]:

                print()
                print("Voice entry stopped.")
                break

            # ------------------------------------------------
            # NEXT
            # ------------------------------------------------

            if text in [
                "next",
                "next row",
                "next cell"
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
                "back",
                "previous",
                "previous row",
                "previous cell"
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
                "clear",
                "delete",
                "clear cell"
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
