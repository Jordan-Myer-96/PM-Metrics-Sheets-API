import statistics
from collections import Counter
from itertools import combinations
from collections import defaultdict
from typing import List, Dict, Tuple, Any
from gspread import Worksheet

from logic.transformations import *

def get_even_odd_data(runs: list[dict]) -> None:
    # Create a dictionary to store the number of wins and total runs for each pickedHat
    picked_hat_stats = {}
    filtered_data = []
    for data_dict in runs:
        if data_dict.get("pickedHat") == "anniv5:EvenOddPack":
            filtered_item = {key: data_dict[key] for key in [ "host","pickedHat", "victory"]}
            filtered_data.append(filtered_item)

    host_stats = defaultdict(lambda: {"total_games": 0, "total_wins": 0})

    # Calculate total games and wins for each host
    for item in filtered_data:
        host = item["host"]
        host_stats[host]["total_games"] += 1
        if item["victory"]:
            host_stats[host]["total_wins"] += 1

    # Calculate win rate for each host and create the output
    output = []
    for host, stats in host_stats.items():
        total_games = stats["total_games"]
        total_wins = stats["total_wins"]
        win_rate = total_wins / total_games if total_games > 0 else 0.0
        output.append([host,total_games,win_rate])
    
    return output

    #         picked_hat = data_dict["pickedHat"]
    #         player_id = data_dict["host"]
    #         victory = data_dict["victory"] 

    #         # Initialize the pickedHat's statistics if not already present
    #         if picked_hat not in picked_hat_stats:
    #             picked_hat_stats[picked_hat] = {"wins": 0, "total_runs": 0}

    #         # Update statistics based on victory
    #         picked_hat_stats[picked_hat]["total_runs"] += 1
    #         if victory:
    #             picked_hat_stats[picked_hat]["wins"] += 1

    # sorted_results = sorted(picked_hat_stats.items(),
    #                         key=lambda x: (x[1]["wins"] / x[1]["total_runs"] if x[1]["total_runs"] > 0 else 0.0),
    #                         reverse=True)

    # hat_win_rate_output = []
    # # Calculate and print the win rate for each pickedHat
    # for picked_hat, stats in sorted_results:
    #     wins = stats["wins"]
    #     total_runs = stats["total_runs"]

    #     hat_win_rate_output.append([del_prefix(picked_hat),wins,total_runs,make_ratio(wins,total_runs)])
    #     #print(f"{del_prefix(picked_hat)}: {make_ratio(wins, total_runs)}")
    # return hat_win_rate_output

