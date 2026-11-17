#Joshua Riley MMSF Code
#Project Start Date: 11/16/26
#Project Finished Date: Work in Progress
#Code Cabailities (currently): Finding MMSF 1, 2, and 3 Standard Cards

#If code isn't running, may need to install openpyxl
#Copy and paste the code below in the command prompt
#When copying and pasting it, make sure it isn't a comment
#pip install openpyxl

from openpyxl import load_workbook
from pathlib import Path

# ==================================================
# MEGA MAN STAR FORCE CARD FINDER
# VERSION 1.0
# CARD CATEGORY: STANDARD
# ==================================================

# Excel workbook location
EXCEL_FILE = Path(__file__).parent / "MMSF Cards.xlsx"

# Currently supported worksheet
SHEET_NAME = "MMSF Standard Cards"

# ==================================================
# GAME INFORMATION
# ==================================================

GAMES = {
    "1": {
        "name": "Mega Man Star Force 1",
        "column": 1
    },

    "2": {
        "name": "Mega Man Star Force 2",
        "column": 4
    },

    "3": {
        "name": "Mega Man Star Force 3",
        "column": 7
    }
}

# ==================================================
# LOAD EXCEL DATABASE
# ==================================================

def load_card_database():

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel workbook not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        data_only=True
    )

    if SHEET_NAME not in workbook.sheetnames:
        workbook.close()

        raise ValueError(
            f"Worksheet '{SHEET_NAME}' not found."
        )

    sheet = workbook[SHEET_NAME]

    database = {}

    for game, information in GAMES.items():

        database[game] = {}

        start_column = information["column"]

        # Read all rows containing card information
        for row in range(3, sheet.max_row + 1):

            card_number = sheet.cell(
                row=row,
                column=start_column
            ).value

            card_name = sheet.cell(
                row=row,
                column=start_column + 1
            ).value

            card_locations = sheet.cell(
                row=row,
                column=start_column + 2
            ).value

            # Skip empty rows
            if card_number is None:
                continue

            # Convert Excel card number to integer
            try:
                card_number = int(card_number)
            except (ValueError, TypeError):
                continue

            # Split multiple locations into individual entries
            if card_locations:

                locations = [
                    location.strip()
                    for location in str(card_locations).splitlines()
                    if location.strip()
                ]

            else:
                locations = []

            # Store card information in memory
            database[game][card_number] = {
                "name": card_name,
                "locations": locations
            }

    workbook.close()

    return database

# ==================================================
# SEARCH FOR MULTIPLE CARDS
# ==================================================

def find_cards(database, game, card_numbers):

    print("\n")
    print("=" * 55)
    print("                  SEARCH RESULTS")
    print("=" * 55)

    print(f"Game: {GAMES[game]['name']}")
    print("Card Category: Standard")
    print(f"Cards Requested: {len(card_numbers)}")

    print("=" * 55)

    found_count = 0
    missing_count = 0

    for card_number in card_numbers:

        card = database[game].get(card_number)

        print(f"\nCARD #{card_number}")
        print("-" * 55)

        if card is None:

            print("Card not found in the database.")

            missing_count += 1

            continue

        found_count += 1

        print(f"Card Name: {card['name']}")

        print("\nPossible Acquisition Locations:")

        if card["locations"]:

            for location in card["locations"]:

                print(f"  - {location}")

        else:

            print("  No locations currently recorded.")

    print("\n")
    print("=" * 55)
    print("                 SEARCH SUMMARY")
    print("=" * 55)

    print(f"Cards successfully found: {found_count}")
    print(f"Cards not found: {missing_count}")

    print("=" * 55)
    print("\n")

# ==================================================
# GET CARD NUMBERS FROM USER
# ==================================================

def get_card_numbers():

    while True:

        card_input = input(
            "\nEnter missing card numbers "
            "(separate with commas): "
        ).strip()

        try:

            if not card_input:
                raise ValueError

            card_numbers = [
                int(number.strip())
                for number in card_input.split(",")
            ]

            if any(number <= 0 for number in card_numbers):
                raise ValueError

            # Remove duplicates while preserving order
            card_numbers = list(dict.fromkeys(card_numbers))

            return card_numbers

        except ValueError:

            print(
                "\nInvalid input. "
                "Enter positive card numbers separated by commas."
            )

# =================================================
# MAIN MENU
# ==================================================

def main():

    print("\n")
    print("=" * 55)
    print("          MEGA MAN STAR FORCE CARD FINDER")
    print("=" * 55)

    print("\nLoading Excel database...")

    try:

        database = load_card_database()

    except Exception as error:

        print(f"\nERROR: {error}")
        input("\nPress Enter to close...")
        return

    print("Database successfully loaded!")

    # Main program loop
    while True:

        print("\n")
        print("=" * 55)
        print("                    MAIN MENU")
        print("=" * 55)

        print("\nSelect a game:\n")

        print("1. Mega Man Star Force 1")
        print("2. Mega Man Star Force 2")
        print("3. Mega Man Star Force 3")

        print("\nQ. Quit")

        game = input("\nEnter selection: ").strip().upper()

        if game == "Q":

            print("\nThank you for using MMSF Card Finder!")
            break

        if game not in GAMES:

            print("\nInvalid selection. Please try again.")
            continue

        print("\nSelected:", GAMES[game]["name"])

        # Standard is the only available card category
        print("Card Category: Standard")

        # Collect missing card numbers
        card_numbers = get_card_numbers()

        # Search database
        find_cards(database, game, card_numbers)

        input("Press Enter to return to the main menu...")

# ==================================================
# START PROGRAM
# ==================================================

if __name__ == "__main__":

    main()
