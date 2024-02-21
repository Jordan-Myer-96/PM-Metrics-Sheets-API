import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread import Worksheet
from logic import insights, custom
from logic.storage import *
import string
import pickle

# Global dictionary to store sublists based on file paths
date_to_metrics = {}
# Dictionairy of packs to cards
pack_to_cards = {}
card_to_pack = {}

spreadsheet_name = "Updated Packmaster Metrics: 07/23-02/24"
# Define the scope and credentials
scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('google_api.json', scope)
print("Got Credentials")
# Authorize
client = gspread.authorize(credentials)
pm_workbook = client.open("Updated Packmaster Metrics: 07/23-02/24")

def merge_cells(sheet,end_col):
    # Send the request to merge cells
    request = {
        
        "mergeCells": {
            "range": {
                "sheetId": sheet.id,
                "startRowIndex": 1,
                "endRowIndex": 2,
                "startColumnIndex": 1,
                "endColumnIndex": end_col
            },
            "mergeType": "MERGE_ROWS"
        }
    }
    return request

def filter_lists_without_colon(input_list):
    return [lst for lst in input_list if not(isinstance(lst[0], str) and ':' in lst[0])]

def clear_and_add(sheet_title, data, *headers, percent_col=None):
    
    try:
        # Try to get the sheet
        sheet = pm_workbook.worksheet(sheet_title)
    except gspread.exceptions.WorksheetNotFound:
        # If the sheet doesn't exist, there's nothing to delete
        sheet = None
    
    if sheet:
        pm_workbook.del_worksheet(sheet)
        print(f"Sheet '{sheet_title}' deleted successfully.")
    
    # Determine the number of rows and columns needed based on the size of the data
    num_rows = len(data) + 1  # Adding 1 for the header row
    num_cols = len(headers)
    
    # Add a new worksheet with the determined number of rows and columns
    new_sheet = pm_workbook.add_worksheet(title=sheet_title, rows=str(num_rows), cols=str(num_cols))
    sheet = pm_workbook.worksheet(sheet_title)

    # Write headers to the first row
    new_sheet.append_row([sheet_title])
    new_sheet.append_row(headers)
    new_sheet.append_rows(data)

    # Apply a filter to all columns
    last_column_letter = chr(ord('A') + num_cols - 1)  # Convert column number to letter
    range_with_headers = f"A2:{last_column_letter}{num_rows + 1}"
    new_sheet.set_basic_filter(range_with_headers)

     # If a percent column is specified, format the column as a percent
    if percent_col:
        percent_column_range = f"{percent_col}2:{percent_col}{num_rows + 1}"
        percent_format = {"numberFormat": {"type": "PERCENT", "pattern": "#,##0.00%"}}
        new_sheet.format(percent_column_range, percent_format)
    
    formats = [
                {
                    "range": "A1:A1",
                    "format": {
                        "textFormat": {
                            "bold": True,
                        },
                        "horizontalAlignment": "CENTER"  # Add center alignment
                    },
                }
            ]
    
    response = sheet.batch_format(formats)
    
    body = {
        "requests": [
            {
                "mergeCells": {
                    "mergeType": "MERGE_ALL",
                    "range": {  # In this sample script, all cells of "A1:C3" of "Sheet1" are merged.
                        "sheetId": sheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_cols
                    }
                }
            }
        ]
    }
    response = pm_workbook.batch_update(body)


    # Send the batch update request
    #response = sheet.batch_update({'requests': batch_requests})



if __name__ == "__main__":

    data_path = os.path.join(os.getcwd(), "data")
    metrics_path = os.path.join(data_path, "metrics")
    print(metrics_path)
    data_file = "data.pkl"
    data_file_path = os.path.join(data_path, data_file)
    print(data_file_path)
    # Check if the data file exists to avoid reprocessing
    if os.path.exists(data_file_path):
        print("Loading Pickle")
        date_to_metrics = load_data_from_pickle(data_file_path)
        print("Loaded Pickle")
        
    else:
        print("Making pickle")
        #date_to_metrics = iterate_directory(metrics_path)
        #print(date_to_metrics.keys())
        #save_data_to_pickle(data_file_path, date_to_metrics)
        date_to_metrics = iterate_directory(metrics_path)
        save_data_to_pickle(data_file_path, date_to_metrics)

    pack_to_cards = load_data_from_json(os.path.join(data_path, "packCards.json"))
    card_to_rarity = load_data_from_json(os.path.join(data_path, "rarities.json"))
    card_to_pack = reverse_and_flatten_dict(pack_to_cards)

    all_data_dict = round_date_keys(date_to_metrics, 0)
    #insights.count_average_win_rate_per_card_split_by_rarity(all_data_dict[""], card_to_pack, card_to_rarity)
    #insights.pack_efficiency_analysis(all_data_dict[""], card_to_pack, card_to_rarity)

#Winrate by Asc
    asc_win_rate_data =  filter_lists_without_colon(insights.count_win_rates(all_data_dict[""]))
    asc_win_rate_headers = ["Ascension","Wins","Total Runs","Winrate"]
    clear_and_add("Ascension Win Rate Test",asc_win_rate_data,*asc_win_rate_headers,percent_col="D")

#Pack victory rate
    pack_victory_data = filter_lists_without_colon(insights.count_pack_victory_rate(all_data_dict[""]))
    pack_victory_headers = ["Pack","Wins","Total Runs","Winrate"]
    clear_and_add("Pack Win Rate",pack_victory_data,*pack_victory_headers,percent_col="D")

#Overall card winrate
    card_win_rate_data = filter_lists_without_colon(insights.count_average_win_rate_per_card(all_data_dict[""], card_to_pack))
    winrate_headers = ["Pack","Card","Wins","Total Runs","Winrate"]
    clear_and_add("Card Win Rate",card_win_rate_data,*winrate_headers,percent_col="E")

# Winrate by card rarity
    rarity_winrate_data = insights.count_average_win_rate_per_card_split_by_rarity(all_data_dict[""], card_to_pack, card_to_rarity)
    rarity_winrate_headers = ["Pack","Card","Wins","Total Runs","Winrate"]
    common_data = filter_lists_without_colon(rarity_winrate_data.get("COMMON", []))
    uncommon_data = filter_lists_without_colon(rarity_winrate_data.get("UNCOMMON", []))
    rare_data = filter_lists_without_colon(rarity_winrate_data.get("RARE", []))

    clear_and_add("Win Rate per card by rarity (Common)",common_data,*rarity_winrate_headers,percent_col="E")
    clear_and_add("Win Rate per card by rarity (Uncommon)",uncommon_data,*rarity_winrate_headers,percent_col="E")
    clear_and_add("Win Rate per card by rarity (Rare)",rare_data,*rarity_winrate_headers,percent_col="E")

#Card Pick Rate
    card_pick_rate_data =  filter_lists_without_colon(insights.count_card_pick_rate(all_data_dict[""], card_to_pack))
    card_pick_rate_headers = ["Pack","Card","Times Picked","Times Offered","Pick Rate"]
    clear_and_add("Card Pick Rate",card_pick_rate_data,*card_pick_rate_headers,percent_col="E")

# Median Winning Deck Size
    median_deck_data = filter_lists_without_colon(insights.count_median_deck_sizes(all_data_dict[""]))
    median_deck_headers = ["Ascension Level","Median Victorious Deck Size"]
    clear_and_add("Median Deck Size by Asc",median_deck_data,*median_deck_headers)

#Hat Pick Rate
    hat_pick_rate_data =  filter_lists_without_colon(insights.count_most_common_picked_hats(all_data_dict[""]))
    hat_pick_rate_headers = ["Hat","Total Runs"]
    clear_and_add("Total Hat Selection",hat_pick_rate_data,*hat_pick_rate_headers)

#Hat Win Rate
    hat_win_rate_data = filter_lists_without_colon(insights.count_win_rate_per_picked_hat(all_data_dict[""]))
    hat_win_rate_headers = ["Hat","Wins","Total Runs","Winrate"]
    clear_and_add("Win Rate per Hat Selected",hat_win_rate_data,*hat_win_rate_headers,percent_col="D")

#Pack Efficiency
    pack_eff_data = filter_lists_without_colon(insights.pack_efficiency_analysis(all_data_dict[""],card_to_pack))
    pack_eff_headers = ["Pack","Wins Containing at least 1 card?","Runs","Frequency Cards Are In Winning Decks"]
    clear_and_add("Pack Efficiency Full",pack_eff_data,*pack_eff_headers,percent_col="D")

