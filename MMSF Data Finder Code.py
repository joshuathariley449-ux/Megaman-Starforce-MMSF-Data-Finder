#Joshua Riley MMSF Code
#Project Start Date: 11/16/26
#Project Finished Date: Work in Progress
#Current capabilities: Finding MMSF 1, 2, and 3 Standard, Mega, and Giga Cards

#If code isn't running, install openpyxl in Command Prompt:
#pip install openpyxl

#Uses the openpyxl package features to enable the excel workbook to be read
from openpyxl import load_workbook
from pathlib import Path

#Excel workbook information
#This is important, this code WILL only run with the excel I've set it up to

#This code makes sure we open the workbook called MMSF Cards.xlsx
#If the name of the excel file were ever to be change, this would need to be changed as well
ExcelWorkbook = Path(__file__).parent / "MMSF Cards.xlsx"

#Within the workbook, there are different sheets of lists
#Depending on the information we want, the program will go through different sheets accordingly

StandardCards = "MMSF Standard Cards" #sheet containing all the standard class cards across all 7 MMSF titles
MegaCards = "MMSF Mega Cards" #sheet containing all the mega class cards across all 7 MMSF titles
GigaCards = "MMSF Giga Cards" #sheet containing all the giga class cards across all 7 MMSF titles

#Game Information
#There are three games technically total in this series (storyline wise)
#for organization, regardless of version I have all the lists below MMSF 1,2, and 3 respectively

#Lines 31-45 create directory to find cards related to MMSF 1, 2, and 3 by which game it is
GameTitles = {
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

#Lines 50-54 create directory to find cards related to MMSF 1, 2, and 3 respectively for Mega Class Cards
#While there are only 3 storyline games, each game (1,2, and 3) has different versions, which also means special version exclusive cards
#In order for this code to work, it has to be able to navigate version exclusive cards, starting with the mega class cards
MegaCards_GameTitles = {
    "1": {"name": "Mega Man Star Force 1", "column": 2},
    "2": {"name": "Mega Man Star Force 2", "column": 5},
    "3": {"name": "Mega Man Star Force 3", "column": 8}
}

#Lines 60-64 create directory to find cards related to MMSF 1, 2, and 3 respectively for Giga Class Cards
#Similar to what I talked about in lines 47-49; there are game version exclusives cards wise for the other games
#All 7 titles have unique giga cards to them, luckily all found the same way
#This little bit of code makes sure the code can navigate the Giga Cards Sheet correctly when finding the card the user requests
GigaCards_GameTitles = {
    "1": {"name": "Mega Man Star Force 1", "column": 2},
    "2": {"name": "Mega Man Star Force 2", "column": 6},
    "3": {"name": "Mega Man Star Force 3", "column": 10}
}

#While there are only 3 games, which lines 27-64 organized into games 1,2, and 3; the versions of those three need to be initialized

#This directory is for Megaman Starforce 1's unique Mega Class Cards
#Starforce one is the only game of the triology that has unique Mega Cards
MegaClassVersions = {
    "1": "Leo",
    "2": "Pegasus",
    "3": "Dragon"
}

#This is for finding the specific Giga Cards unique to the seven different titles of MMSF
GigaClassVersions = {
    "1": {
        "1": "Leo",
        "2": "Pegasus",
        "3": "Dragon"
    },
    "2": {
        "1": "Zerker",
        "2": "Ninja",
        "3": "Saurian"
    },
    "3": {
        "1": "Black Ace",
        "2": "Red Joker"
    }
}

#These next lines are the actual code that look for the said information outlined from lines 16-92
#lines 96-110 are how the workbook is loaded, and if it isn't loaded, the error message it will show
def open_worksheet(sheet_name):
    if not ExcelWorkbook.exists():
        raise FileNotFoundError(
            f"Excel workbook not found: {ExcelWorkbook}"
        )
    
    workbook = load_workbook(ExcelWorkbook, data_only=True)
    
    if sheet_name not in workbook.sheetnames:
        workbook.close()
        raise ValueError(
            f"Worksheet '{sheet_name}' not found in the Excel workbook."
        )

    return workbook, workbook[sheet_name]

#Code realated to loading the Standard Class Cards Information
def load_card_database():
    workbook, sheet = open_worksheet(StandardCards)
    database = {}

    for game, information in GameTitles.items():
        database[game] = {}
        start_column = information["column"]

        for row in range(3, sheet.max_row + 1):
            card_number = sheet.cell(row=row, column=start_column).value
            card_name = sheet.cell(row=row, column=start_column + 1).value
            card_locations = sheet.cell(row=row, column=start_column + 2).value

            if card_number is None:
                continue

            try:
                card_number = int(card_number)
            except (ValueError, TypeError):
                continue

            locations = (
                [
                    location.strip()
                    for location in str(card_locations).splitlines()
                    if location.strip()
                ]
                if card_locations else []
            )

            database[game][card_number] = {
                "name": card_name,
                "locations": locations
            }

    workbook.close()
    return database

#Code realated to loading the Mega Class Cards Information
def load_mega_card_database():
    workbook, sheet = open_worksheet(MegaCards)
    database = {game: {} for game in GameTitles}

    for game in GameTitles:
        start_column = MegaCards_GameTitles[game]["column"]

        for row in range(3, sheet.max_row + 1):
            # The existing Mega worksheet stores a version in column A
            # for Star Force 1. Star Force 2 and 3 are not version-split
            # in the current Mega database structure.
            if game == "1":
                game_version = sheet.cell(row=row, column=1).value
            else:
                game_version = None

            card_number = sheet.cell(row=row, column=start_column).value
            card_name = sheet.cell(row=row, column=start_column + 1).value
            card_locations = sheet.cell(row=row, column=start_column + 2).value

            if card_number is None:
                continue

            try:
                card_number = int(card_number)
            except (ValueError, TypeError):
                continue

            locations = (
                [
                    location.strip()
                    for location in str(card_locations).splitlines()
                    if location.strip()
                ]
                if card_locations else []
            )

            card_information = {
                "name": card_name,
                "locations": locations
            }

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

                database[game][card_number][version_key] = card_information

            else:
                database[game][card_number] = card_information

    workbook.close()
    return database

#Code realated to loading the Giga Class Cards Information
def load_giga_card_database():
    workbook, sheet = open_worksheet(GigaCards)

    #Each game has its own card-number dictionary.
    #Each card number then contains entries separated by game version.
    database = {game: {} for game in GameTitles}

    for game in GameTitles:
        start_column = GigaCards_GameTitles[game]["column"]
        version_column = start_column - 1

        for row in range(3, sheet.max_row + 1):
            game_version = sheet.cell(
                row=row,
                column=version_column
            ).value

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

            if card_number is None or game_version is None:
                continue

            try:
                card_number = int(card_number)
            except (ValueError, TypeError):
                continue

            version_text = str(game_version).strip().lower()

            #Match the worksheet's version text to the game's
            #supported version names
            version_key = None
            for version_name in GigaClassVersions[game].values():
                if version_text == version_name.lower():
                    version_key = version_name
                    break

            if version_key is None:
                continue

            locations = (
                [
                    location.strip()
                    for location in str(card_locations).splitlines()
                    if location.strip()
                ]
                if card_locations else []
            )

            if card_number not in database[game]:
                database[game][card_number] = {}

            database[game][card_number][version_key] = {
                "name": card_name,
                "locations": locations
            }

    workbook.close()
    return database

#Code to search for the Standard Cards, if requested by the user
def find_standard_cards(database, game, card_numbers):
    print("\n")
    print("=" * 55)
    print("                  SEARCH RESULTS")
    print("=" * 55)
    print(f"Game: {GameTitles[game]['name']}")
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

#Code to search for the Mega Cards, if requested by the user
def find_mega_cards(database, game, card_numbers, version=None):
    print("\n")
    print("=" * 55)
    print("                  SEARCH RESULTS")
    print("=" * 55)
    print(f"Game: {GameTitles[game]['name']}")
    print("Card Category: Mega")

    if game == "1":
        print(f"Game Version: {version}")

    print(f"Cards Requested: {len(card_numbers)}")
    print("=" * 55)

    found_count = 0
    missing_count = 0

    for card_number in card_numbers:
        card = database[game].get(card_number)

        # Preserve the existing Mega behavior:
        # Star Force 1 checks the selected version, then shared cards.
        if game == "1" and card is not None:
            selected_card = card.get(version)

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

#Code to search for the Giga Cards, if requested by the user
def find_giga_cards(database, game, card_numbers, version):
    print("\n")
    print("=" * 55)
    print("                  SEARCH RESULTS")
    print("=" * 55)
    print(f"Game: {GameTitles[game]['name']}")
    print("Card Category: Giga")
    print(f"Game Version: {version}")
    print(f"Cards Requested: {len(card_numbers)}")
    print("=" * 55)

    found_count = 0
    missing_count = 0

    for card_number in card_numbers:
        version_cards = database[game].get(card_number, {})
        card = version_cards.get(version)

        print(f"\nCARD #{card_number}")
        print("-" * 55)

        if card is None:
            print(
                f"Card not found for the selected {version} version."
            )
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

#In order to search for the cards you want, the user needs to provide the number the card is correlated to
#Standard Class Cards are numbered 1-150
#Mega Class Cards are numbered the following for each game:
#MMSF 1: 1-30 
#MMSF 2: 1-42
#MMSF 3: 1-45
#Giga Class Cards are numbered 1-5
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

            # Remove duplicates while preserving order.
            card_numbers = list(dict.fromkeys(card_numbers))
            return card_numbers

        except ValueError:
            print(
                "\nInvalid input. "
                "Enter positive card numbers separated by commas."
            )

#This is the code that allows the user to define which version of Megaman Starforce 1 they are playing
#This is needed so the proper Mega Class Cards are obtained from each version
def select_version():
    while True:
        print("\n")
        print("=" * 55)
        print("             SELECT GAME VERSION")
        print("=" * 55)
        print("\n1. Leo")
        print("2. Pegasus")
        print("3. Dragon")

        selection = input("\nEnter selection: ").strip()

        if selection in MegaClassVersions:
            return MegaClassVersions[selection]

        print("\nInvalid selection. Please try again.")

#This is the code that allows the user to define which version of Megaman Starforce 1,2, and 3 they are playing
#This is needed so the proper Giga Class Cards are obtained from each version
def select_giga_version(game):
    available_versions = GigaClassVersions[game]

    while True:
        print("\n")
        print("=" * 55)
        print("             SELECT GIGA GAME VERSION")
        print("=" * 55)

        for number, version_name in available_versions.items():
            print(f"{number}. {version_name}")

        selection = input("\nEnter selection: ").strip()

        if selection in available_versions:
            return available_versions[selection]

        print("\nInvalid selection. Please try again.")

#This is the code for the main program
#In other words this the code that uses all the lines before it to ask the user for the information to find all wanted cards
def main():
    print("\n")
    print("=" * 55)
    print("          MEGA MAN STAR FORCE CARD FINDER")
    print("=" * 55)
    print("\nLoading Excel databases...")

    try:
        standard_database = load_card_database()
        mega_database = load_mega_card_database()
        giga_database = load_giga_card_database()

    except Exception as error:
        print(f"\nERROR: {error}")
        input("\nPress Enter to close...")
        return

    print("Standard database successfully loaded!")
    print("Mega database successfully loaded!")
    print("Giga database successfully loaded!")

    #Main loop that gets the main program to work
    while True:
        print("\n")
        print("=" * 55)
        print("                    MAIN MENU")
        print("=" * 55)
        print("\nSelect card category:\n")
        print("1. Standard Class")
        print("2. Mega Class")
        print("3. Giga Class")
        print("\nQ. Quit")

        category = input("\nEnter selection: ").strip().upper()

        if category == "Q":
            print("\nThank you for using MMSF Card Finder!")
            break

        if category not in ["1", "2", "3"]:
            print("\nInvalid selection. Please try again.")
            continue
        
        #Where the game version is selected, uses a 1,2 and 3 format
        print("\n")
        print("=" * 55)
        print("                  SELECT GAME")
        print("=" * 55)
        print("\n1. Mega Man Star Force 1")
        print("2. Mega Man Star Force 2")
        print("3. Mega Man Star Force 3")

        game = input("\nEnter selection: ").strip()

        if game not in GameTitles:
            print("\nInvalid selection. Returning to main menu.")
            continue

        print("\nSelected:", GameTitles[game]["name"])

        #Loop for finding standard class cards across the 7 titles
        if category == "1":
            print("Card Category: Standard")
            card_numbers = get_card_numbers()

            find_standard_cards(
                standard_database,
                game,
                card_numbers
            )

        #Loop for finding mega class cards across the 7 titles
        elif category == "2":
            print("Card Category: Mega")
            version = None

            # Preserve the existing Mega version selection.
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

        #Loop for finding giga class cards across the 7 titles
        elif category == "3":
            print("Card Category: Giga")

            # Giga cards are version-specific in all three GameTitles.
            version = select_giga_version(game)
            print("\nSelected Version:", version)

            card_numbers = get_card_numbers()

            find_giga_cards(
                giga_database,
                game,
                card_numbers,
                version
            )

        input("Press Enter to return to the main menu...")

#This is to start the program
if __name__ == "__main__":
    main()
