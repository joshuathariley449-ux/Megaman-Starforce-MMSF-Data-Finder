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
# EXCEL WORKBOOK INFORMATION
# ==================================================

EXCEL_FILE = Path(__file__).parent / "MMSF Cards.xlsx"

STANDARD_SHEET = "MMSF Standard Cards"
MEGA_SHEET = "MMSF Mega Cards"

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
# MEGA CARD COLUMN INFORMATION
# ==================================================

MEGA_GAMES = {
    "1": {
        "name": "Mega Man Star Force 1",
        "column": 2
    },

    "2": {
        "name": "Mega Man Star Force 2",
        "column": 5
    },

    "3": {
        "name": "Mega Man Star Force 3",
        "column": 8
    }
}
# ==================================================
# STAR FORCE 1 VERSION INFORMATION
# ==================================================

VERSIONS = {
    "1": "Leo",
    "2": "Pegasus",
    "3": "Dragon"
}

# ==================================================
# LOAD STANDARD CARD DATABASE
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

    if STANDARD_SHEET not in workbook.sheetnames:
        workbook.close()

        raise ValueError(
            f"Worksheet '{STANDARD_SHEET}' not found."
        )

    sheet = workbook[STANDARD_SHEET]

    database = {}

    for game, information in GAMES.items():

        database[game] = {}

        start_column = information["column"]

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

            if card_number is None:
                continue

            try:
                card_number = int(card_number)
            except (ValueError, TypeError):
                continue

            if card_locations:

                locations = [
                    location.strip()
                    for location in str(card_locations).splitlines()
                    if location.strip()
                ]

            else:
                locations = []

            database[game][card_number] = {
                "name": card_name,
                "locations": locations
            }

    workbook.close()

    return database

# ==================================================
# LOAD MEGA CARD DATABASE
# ==================================================

def load_mega_card_database():

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel workbook not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        data_only=True
    )

    if MEGA_SHEET not in workbook.sheetnames:
        workbook.close()

        raise ValueError(
            f"Worksheet '{MEGA_SHEET}' not found."
        )

    sheet = workbook[MEGA_SHEET]

    database = {}

    for game, information in GAMES.items():

        database[game] = {}

        start_column = MEGA_GAMES[game]["column"]

        for row in range(3, sheet.max_row + 1):

            # Star Force 1 has an additional version column
            if game == "1":

                game_version = sheet.cell(
                    row=row,
                    column=1
                ).value

            else:

                game_version = None

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

            # Skip empty card numbers
            if card_number is None:
                continue

            try:
                card_number = int(card_number)
            except (ValueError, TypeError):
                continue

            # Split locations into individual entries
            if card_locations:

                locations = [
                    location.strip()
                    for location in str(card_locations).splitlines()
                    if location.strip()
                ]

            else:

                locations = []

            card_information = {
                "name": card_name,
                "locations": locations
            }

            # ==========================================
            # STAR FORCE 1 VERSION HANDLING
            # ==========================================

            if game == "1":

                if game_version is None:
                    continue

                version_text = str(game_version).strip().lower()

                if "leo" in version_text and "pegasus" in version_text:

                    version_key = "All"

                elif version_text == "leo":

                    version_key = "Leo"

                elif version_text == "pegasus":

                    version_key = "Pegasus"

                elif version_text == "dragon":

                    version_key = "Dragon"

                else:

                    continue

                if card_number not in database[game]:

                    database[game][card_number] = {}

                database[game][card_number][version_key] = (
                    card_information
                )

            # ==========================================
            # STAR FORCE 2 AND 3
            # ==========================================

            else:

                database[game][card_number] = (
                    card_information
                )

    workbook.close()

    return database

# ==================================================
# SEARCH STANDARD CARDS
# ==================================================

def find_standard_cards(database, game, card_numbers):

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
# SEARCH MEGA CARDS
# ==================================================

def find_mega_cards(database, game, card_numbers, version=None):

    print("\n")
    print("=" * 55)
    print("                  SEARCH RESULTS")
    print("=" * 55)

    print(f"Game: {GAMES[game]['name']}")
    print("Card Category: Mega")

    if game == "1":
        print(f"Game Version: {version}")

    print(f"Cards Requested: {len(card_numbers)}")

    print("=" * 55)

    found_count = 0
    missing_count = 0

    for card_number in card_numbers:

        card = database[game].get(card_number)

        # ==========================================
        # SELECT CORRECT STAR FORCE 1 VERSION
        # ==========================================

        if game == "1" and card is not None:

            # First look for the selected version
            selected_card = card.get(version)

            # If unavailable, check for shared cards
            if selected_card is None:

                selected_card = card.get("All")

            card = selected_card

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

# ==================================================
# SELECT STAR FORCE 1 VERSION
# ==================================================

def select_version():

    while True:

        print("\n")
        print("=" * 55)
        print("             SELECT GAME VERSION")
        print("=" * 55)

        print("\n1. Leo")
        print("2. Pegasus")
        print("3. Dragon")

        selection = input(
            "\nEnter selection: "
        ).strip()

        if selection in VERSIONS:

            return VERSIONS[selection]

        print("\nInvalid selection. Please try again.")

# ==================================================
# MAIN PROGRAM
# ==================================================

def main():

    print("\n")
    print("=" * 55)
    print("          MEGA MAN STAR FORCE CARD FINDER")
    print("=" * 55)

    print("\nLoading Excel databases...")

    try:

        standard_database = load_card_database()

        mega_database = load_mega_card_database()

    except Exception as error:

        print(f"\nERROR: {error}")
        input("\nPress Enter to close...")
        return

    print("Standard database successfully loaded!")
    print("Mega database successfully loaded!")

    # ==============================================
    # MAIN PROGRAM LOOP
    # ==============================================

    while True:

        print("\n")
        print("=" * 55)
        print("                    MAIN MENU")
        print("=" * 55)

        print("\nSelect card category:\n")

        print("1. Standard Class")
        print("2. Mega Class")

        print("\nQ. Quit")

        category = input(
            "\nEnter selection: "
        ).strip().upper()

        if category == "Q":

            print("\nThank you for using MMSF Card Finder!")
            break

        if category not in ["1", "2"]:

            print("\nInvalid selection. Please try again.")
            continue

        # ==========================================
        # SELECT GAME
        # ==========================================

        print("\n")
        print("=" * 55)
        print("                  SELECT GAME")
        print("=" * 55)

        print("\n1. Mega Man Star Force 1")
        print("2. Mega Man Star Force 2")
        print("3. Mega Man Star Force 3")

        game = input(
            "\nEnter selection: "
        ).strip()

        if game not in GAMES:

            print("\nInvalid selection. Returning to main menu.")
            continue

        print("\nSelected:", GAMES[game]["name"])

        # ==========================================
        # STANDARD CLASS SEARCH
        # ==========================================

        if category == "1":

            print("Card Category: Standard")

            card_numbers = get_card_numbers()

            find_standard_cards(
                standard_database,
                game,
                card_numbers
            )

        # ==========================================
        # MEGA CLASS SEARCH
        # ==========================================

        elif category == "2":

            print("Card Category: Mega")

            version = None

            # Only Star Force 1 requires version selection
            if game == "1":

                version = select_version()

                print("\nSelected Version:", version)

            card_numbers = get_card_numbers()

            find_mega_cards(
                mega_database,
                game,
                card_numbers,
                version
            )

        input("Press Enter to return to the main menu...")

# ==================================================
# START PROGRAM
# ==================================================

if __name__ == "__main__":

    main()
Q
