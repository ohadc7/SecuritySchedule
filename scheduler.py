#!/proj/mislcad/areas/DAtools/tools/python/3.10.1/bin/python3

# The algorithm is based on TTR
# TTR == Time to rest

# FIXME: debug fairness
# ./scheduler.py Schedule_2.xlsx --prev "2023-10-25" --next "2023-10-26" --d 30 --pos 2 --seed 1 | tee log_2
# See Benzi

# FIXME: errors should be warnings by default, error for developer (controlled with --arg)

# Nadav:
# #######
# Analyze fairness (add average, more weight to nights, OR, better, night start, day stars, two separate columns)
# Improve verify() function - currently checks that TTR is observed, check also that now two nights in a row
# Add randomization for get_lowest_ttr() - randomize people with the same TTR - check if improves fairness
#
# Consider:
# - Weight per position, per hour
# - Read from XLS, same as team_size or action
# - Use weight instead of TTR in set_ttr()
# Flusk? Pygame?
# Make object-oriented:
#   Class ttr_db
#   Class schedule
#   Other?

# Later
# - Consider dying output by fairness, manual fix
# - Consider: MAX_TTR = MIN_TTR * 2 - 1
# - Consider: post-processing to fix fairness
# - Consider making the TTR part of the CFG (add column)
# - Another way - add column 'weight' fo position CFG
#   Can split position into positions (night, day), each with different weight
# - Allow planning for hour range (can help planning for Shabbat)
# - Support list of people per position
# - Add personal constraints
# - Improve XLS parsing (read once, all sheets, then parse)
# - Run fairness check on range of sheets, possibly without generating

import sys
import os
import random
import pandas as pd
import argparse
import math
import copy

# For writing XLS file
import openpyxl
from openpyxl.styles               import PatternFill
from openpyxl.utils.units          import cm_to_EMU
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.styles               import Font
from datetime                      import datetime, timedelta

##################################################################################
# Constants
##################################################################################

HOURS_IN_DAY     = 24
NUM_OF_POSITIONS = 5
COLUMN_WIDTH     = 27
LINE_WIDTH       = 10 + NUM_OF_POSITIONS*COLUMN_WIDTH

# FIXME: either remove or add explanation.
# Currently fails when using unique list, needs debug
night_hours_rd   = [1, 2, 3, 4]
night_hours_wr   = [23, 0, 1, 2, 3, 4]

##################################################################################
# Enums
##################################################################################

# Actions
NONE = 0; SWAP = 1; RESIZE = 2
# Colors
PINK = 0; BLUE = 1; GREEN  = 2; YELLOW = 3; PURPLE = 4

##################################################################################
# Configurable default values
##################################################################################

SEED         = 1
DAYS_TO_PLAN = 1
TTR_NIGHT    = 9
TTR_DAY      = 4

##################################################################################
# Utils
##################################################################################

def print_delimiter(): print("#" * LINE_WIDTH)
def error(message):    print('Error: '+message); exit(1)
def warning(message):  print('Warning: '+message);

##################################################################################
# Functions
##################################################################################
def parse_arguments():
    # Create an ArgumentParser object
    parser = argparse.ArgumentParser(description="Command-line parser")

    # Define the command-line arguments
    parser.add_argument("file_name",   type=str,            help="XLS file name ")
    parser.add_argument("--seed",      type=int,            help="Seed")
    parser.add_argument("--prev",      type=str,            help="Prev schedule sheet name (optional, default is today's date)")
    parser.add_argument("--next",      type=str,            help="Next schedule sheet name (optional, default is tomorrow's date)")
    parser.add_argument("--write",     action="store_true", help="Do write result to the XLS file")
    parser.add_argument("--days",      type=int,            help="Number of days to schedule")
    parser.add_argument("--positions", type=int,            help="Number of positions")
    parser.add_argument("--ttrn",      type=int,            help="Minimum time to rest after NIGHT shift")
    parser.add_argument("--ttrd",      type=int,            help="Minimum time to rest after DAY shift")

    # Parse the command-line arguments
    args = parser.parse_args()

    # Access the parsed arguments
    xls_file_name = args.file_name
    do_write      = args.write

    # Configure global variables
    if args.days:      global DAYS_TO_PLAN;     DAYS_TO_PLAN     = args.days
    if args.positions: global NUM_OF_POSITIONS; NUM_OF_POSITIONS = args.positions
    if args.seed:      global SEED;             SEED             = args.seed; random.seed(SEED)
    if args.ttrn:      global TTR_NIGHT;        TTR_NIGHT        = args.ttrn
    if args.ttrd:      global TTR_DAY;          TTR_DAY          = args.ttrd

    # Sanity checks
    if not os.path.exists(xls_file_name):                  error(f"File {xls_file_name} does not exist.")
    if do_write and not os.access(xls_file_name, os.W_OK): error(f"File {xls_file_name} is not writable.")

    return args.prev, args.next, xls_file_name, args.write

##################################################################################
# Build DB from "List of people"
# Key:   name
# Value: remaining time to rest (TTR) - set to 0
def build_people_db(xls_file_name):
    # Get list of names from XLS
    names = extract_column_from_sheet(xls_file_name, "List of people", "People")

    # Build people DB
    db = {}
    for name in names:
        if type(name) is str:
            db[name[::-1]] = 0

    print(f"Found {len(db.keys())} people in List of people: {db.keys()}")
    return db

##################################################################################
# Extract column from sheet
def extract_column_from_sheet(xls_file_name, sheet_name, column_name):
    df = pd.read_excel(xls_file_name, sheet_name=sheet_name)

    # Check if the column exists in the DataFrame
    if column_name in df.columns:
        # Access and print the contents of the column
        column_list = df[column_name].tolist()
        return column_list
    else:
        error(f"Column '{column_name}' not found in '{sheet_name}'.")

##################################################################################
# Get configurations
# Consider:
#    CFG[position_index][hour]{action}        = action
#    CFG[position_index][hour]{team_size}     = team_size
#    CFG[position_index][hour]{position_name} = position_name (for debug/error messages)
#    CFG[position_index][hour]{weight}        = weight
def get_cfg(xls_file_name):
    cfg_action = []
    for position in range(NUM_OF_POSITIONS):
        sheet_name = "Position "+str(position+1)
        position_cfg_action = extract_column_from_sheet(xls_file_name, sheet_name, "Action")
        if len(position_cfg_action) != HOURS_IN_DAY:
            error("In sheet "+sheet_name+", swap list unexpected length: " + len(position_cfg_action));
        cfg_action.append(position_cfg_action)

    cfg_team_size = []
    for position in range(NUM_OF_POSITIONS):
        sheet_name = "Position "+str(position+1)
        position_cfg_team_size = extract_column_from_sheet(xls_file_name, sheet_name, "Team size")
        if len(position_cfg_team_size) != HOURS_IN_DAY:
            error("In sheet "+sheet_name+", team size list unexpected length: " + len(position_cfg_team_size));
        position_cfg_team_size_int = []
        for member in position_cfg_team_size:
            position_cfg_team_size_int.append(int(member))
        cfg_team_size.append(position_cfg_team_size)

    cfg_position_name = []
    for position in range(NUM_OF_POSITIONS):
        sheet_name = "Position "+str(position+1)
        names = extract_column_from_sheet(xls_file_name, sheet_name, "Name")
        cfg_position_name.append(names[0][::-1])

    return cfg_action, cfg_team_size, cfg_position_name

##################################################################################
# Get the previous schedule
def get_prev_schedule(xls_file_name, sheet_name, cfg_position_names):
    prev_schedule = []
    if not sheet_name:
        sheet_name = str(datetime.date.today())

    position_teams = []
    for position in range(NUM_OF_POSITIONS):
        position_name = cfg_position_names[position][::-1]
        position_teams.append(extract_column_from_sheet(xls_file_name, sheet_name, position_name))

    for hour in range(HOURS_IN_DAY):
        prev_schedule.append([])
        for team in position_teams:
            team_str = str(team[hour])[::-1]
            if team_str == 'nan':
                team_list = []
            else:
                team_list = team_str.split(",")
            prev_schedule[hour].append(team_list)

    return prev_schedule

##################################################################################
# Check for swap - get string, return bool
def get_action(cfg_action, hour, team_size):
    swap_str = str(cfg_action[hour])

    # Check if need to swap
    if swap_str == 'swap':
        return SWAP
    elif swap_str == 'resize':
        return RESIZE
    elif swap_str == 'nan':
        return NONE
    else:
        error('Unrecognized text ' + swap_str)

    return do_swap

##################################################################################
# Choose team
def choose_team(hour, night_list, ttr_db, team_size):
    #print(f"Choose team for hour {hour}\nNight_list: {night_list}")# \nDB: {ttr_db}")
    team = []
    if hour in night_hours_wr: is_night = 1
    else:                      is_night = 0

    for i in range(team_size):
        name = get_lowest_ttr(ttr_db)
        if is_night:
            # Using for (instead of while)to avoid endless loop
            # Possibly no choice but to take from night watchers
            for i in range(len(ttr_db)):
                if name not in night_list:
                    break
                name = get_lowest_ttr(ttr_db, i+1)

            # Check fairness
            if ttr_db[name] > 0:
                error(f"Chosen {name} with TTR {ttr_db[name]}\nNight list: {night_list}\nSorted: {dict(sorted(ttr_db.items(), key=lambda item: item[1]))}")
            if name in night_list:
                error(f"At {hour}:00, must take night watcher")

        set_ttr(hour, name, ttr_db)
        team.append(name)

    #print(f"Chosen team: {team}")

    return team

##################################################################################
# Resize team
def resize_team(hour, night_list, ttr_db, old_team, new_team_size):

    if new_team_size == 0:
        return [""]

    # Create new team list (to avoid modifying the previous hour value, team is passed by reference)
    new_team = old_team.copy()

    old_team_size = len(old_team)
    if old_team_size == new_team_size:
        error(f"Resize at {hour}:00: old_team_size == new_team_size == {old_team_size}")

    # Resize
    if new_team_size < old_team_size:
        # Reduce team size
        for i in range(old_team_size-new_team_size):
            random_index = random.randint(0, old_team_size-1)
            released = new_team.pop(random_index)
    else:
        # Increase team size
        new_team = choose_team(hour, night_list, ttr_db, new_team_size-old_team_size)

    return new_team

##################################################################################
# Make the assignments
def build_schedule(prev_schedule, night_list, ttr_db, cfg_action, cfg_team_size):

    schedule = [[] for _ in range(HOURS_IN_DAY)]
    # Stores the current team at the specific position
    # If no action, the same team continues to the next hour
    teams    = [[] for _ in range(NUM_OF_POSITIONS)]

    for hour in range(HOURS_IN_DAY):
        # Choose teams
        for position in range(NUM_OF_POSITIONS):
            team_size = cfg_team_size[position][hour]
            action    = get_action(cfg_action[position], hour, team_size)
            team      = teams[position]

            if action == SWAP:
                team = choose_team(hour, night_list, ttr_db, team_size)
            elif action == RESIZE:
                team = resize_team(hour, night_list, ttr_db, team, team_size)
            elif hour == 0:
                team = prev_schedule[HOURS_IN_DAY-1][position]
                # Note: these people should be recorded as night watchers
                # They are not on the list, because they started the shift at "day hours" (23:00)
                for name in team: night_list.append(name)

            # Put the team in the schedule
            schedule[hour].append(team)
            teams[position] = team

            # Even if there was no swap, the chosen team should get its TTS
            for name in team: set_ttr(hour, name, ttr_db)

        # Update TTS (per hour)
        decrement_ttr(ttr_db)

    return schedule
##################################################################################
# DB utils

# Set TTR for name
def set_ttr(hour, name, db):
    if hour in night_hours_rd: db[name] = TTR_NIGHT+1
    else:                      db[name] = TTR_DAY+1

# For each person, decrement the remaining "time to rest"
def decrement_ttr(db):
    for name in db:
        db[name] -= 1

# Get available people from DB (with TTR == 0)
def get_available(db):
    available = []
    for name, ttr in db.items():
        if ttr == 0:
            available.append(name)
    return available

# Print DB
def print_db(header, db):
    print(header)
    for name in db:
        print(f"{name.ljust(COLUMN_WIDTH)}{db[name]}")

# Get the name with lowest TTR value
# Offset allows to skip N lowest values
def get_lowest_ttr(db, offset=0):

    # Sort the dictionary by values in ascending order
    # To sort in descending order, add `, reverse=True` to the sorted function
    # FIXME: shuffle the names with the same value
    sorted_db = dict(sorted(db.items(), key=lambda item: item[1]))
    #print(f"Sorted DB: {sorted_db}")

    # Create an iterator over the dictionary items
    iter_items = iter(sorted_db.items())

    # Skip offset
    for i in range(offset+1): item = next(iter_items)

    # Get name
    name = item[0]
    return name

##################################################################################
# Update TTR DB with previous schedule
# Result: DB, list
def update_db_with_prev_schedule(valid_names, db, schedule):
    night_list = []

    for hour in range(HOURS_IN_DAY):
        is_night = 1 if hour in night_hours_rd else 0

        for position in range(NUM_OF_POSITIONS):
            team = schedule[hour][position]

            for name in team:
                # Ignore people that are not on the list
                if not name in valid_names: continue

                # Note: "+1" is needed to cancel the following decrement of the whole DB
                set_ttr(hour, name, db)
                if is_night:
                    if name not in night_list: night_list.append(name)

        # Update TTS (for each hour, not for each position)
        decrement_ttr(db)

    return night_list

##################################################################################
# Print schedule
def print_schedule(schedule, cfg_position_name):
    print_delimiter()
    header = "Hour\t"
    for position in range(NUM_OF_POSITIONS):
        header += (cfg_position_name[position]).ljust(COLUMN_WIDTH)+"\t"
    print(header)
    print_delimiter()

    for hour in range(HOURS_IN_DAY):
        if hour >= len(schedule):
            error("No schedule for hour "+"{:02d}:00".format(hour))
        line_str = "{:02d}:00\t".format(hour)
        for team in schedule[hour]:
            line_str += ",".join(team).ljust(COLUMN_WIDTH)
            line_str += "\t"
        print(line_str)

##################################################################################
# Write schedule to XLS file
def write_schedule_to_xls(xls_file_name, schedule, sheet_name, cfg_position_name):
    # Open an existing Excel file
    workbook = openpyxl.load_workbook(xls_file_name)

    # Get sheet name for output (only if not provided by the user)
    if not sheet_name:
        sheet_name = str(datetime.date.today() + datetime.timedelta(days=1))

    # Create a new worksheet
    worksheet = workbook.create_sheet(title=sheet_name)

    # Set column width
    worksheet.sheet_format.defaultColWidth = 30

    # Build header row
    header_row = ["Time"]
    for name in cfg_position_name:
        header_row.append(name[::-1])
    worksheet.append(header_row)

    # Write the data to the worksheet
    for hour in range(HOURS_IN_DAY):
        row = schedule[hour]
        row_of_str = ["{:02d}:00\t".format(hour)]
        for team in row:
            row_of_str.append((",".join(team))[::-1])
        worksheet.append(row_of_str)

    # Add colors
    color_worksheet(worksheet)

    # Save the workbook to a file
    workbook.save(xls_file_name)

##################################################################################
# Color the worksheet
def color_worksheet(worksheet):
    color_column(worksheet, 2, PINK)
    color_column(worksheet, 3, GREEN)
    color_column(worksheet, 4, YELLOW)
    color_column(worksheet, 5, BLUE)
    color_column(worksheet, 6, PURPLE)

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

##################################################################################
# Verify result
def color_column(worksheet, index, color):
    # Create a PatternFill object with the required color
    if color   == PINK:
        fill       = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
    elif color == BLUE:
        fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    elif color == GREEN:
        fill = PatternFill(start_color="98FB98", end_color="98FB98", fill_type="solid")
    elif color == YELLOW:
        fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
    elif color == PURPLE:
        fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
    else:
        error(f"Undefined color {color}")

    # Get data from the worksheet
    data = []
    for row in worksheet.iter_rows(values_only=True):
        data.append(list(row))

    for row_num, row in enumerate(data, 1):
        for col_num, value in enumerate(row, 1):
            cell = worksheet.cell(row=row_num, column=col_num, value=value)
            if col_num == index:  # Check if it's the first column
                cell.fill = fill

##################################################################################
# Check fairness
def check_fairness(db, schedule):

    # Init hours_served
    hours_served = {}
    for name in db:
        hours_served[name] = 0

    # Calculate hours_served
    for hour in range(len(schedule)):
        for team in schedule[hour]:
            for name in team:
                if name != 'nan' and name in hours_served:
                    hours_served[name] += 1

    # Report
    print_delimiter()
    print(f"Check fairness")
    print_delimiter()
    for name in hours_served:
        print(f"Name: {name.ljust(COLUMN_WIDTH)} served: {str(hours_served[name]).ljust(4)}\t"+("*"*hours_served[name]))

    # Calculate average
    total = sum(value for value in hours_served.values())
    average = int(total / len(hours_served))

    # Print average
    print_delimiter()
    print(f"Average: {str(average).ljust(COLUMN_WIDTH-3)} served: {str(average).ljust(4)}\t" + ("*" * average))
    print_delimiter()

    return 1

##################################################################################
# Verify result
def verify(db, schedule):
    print_delimiter()
    print(f"Verify total ({len(schedule)} lines)")

    # Init last_served
    last_served = {}
    for name in db:
        last_served[name] = -1

    # Check schedule
    for hour in range(len(schedule)):
        line = schedule[hour]
        #print(f"Verify ({hour}): {line}")
        for team in line:
            for name in team:
                # Ignore people that were removed from the list
                if name in last_served:
                    last_served_hour = last_served[name]
                    if last_served_hour != -1:
                        diff = hour - last_served_hour - 1
                        expected_ttr = TTR_NIGHT if last_served_hour in night_hours_rd else TTR_DAY
                        if diff < expected_ttr and diff > 0:
                            error(f"Poor {name} did not get his {expected_ttr} hour rest (served at {last_served_hour}, then at {hour})")
                    last_served[name] = hour

##################################################################################
# Check who wasn't assigned
def check_for_idle(db, schedule):

    # Init participated
    participated = {}
    for name in db:
        participated[name] = 0

    # Collect data from schedule
    for hour in range(len(schedule)):
        line = schedule[hour]
        for team in line:
            for name in team:
                participated[name] = 1

    # Check who didn't participate
    not_assigned = [item for item in participated.keys() if participated[item] == 0]
    if not_assigned:
        print_delimiter()
        print(f"Not assigned: {not_assigned}")

    return

##################################################################################
# Get next date, based on previous date
def get_next_date(prev_date_str):
    # Specify the format of the date string
    date_format = "%Y-%m-%d"

    # Parse the date string into a datetime object
    prev_obj = datetime.strptime(prev_date_str, date_format)
    next_obj = prev_obj + timedelta(days=1)
    next_date_str = next_obj.strftime(date_format)

    return next_date_str

##################################################################################
# Main
##################################################################################
def main():

    # Parse script arguments
    prev_name, next_name, xls_file_name, do_write = parse_arguments()

    # Build TTR DB {name} -> {time to rest}
    ttr_db = build_people_db(xls_file_name)
    print(f"Orig DB length = {len(ttr_db)}")
    valid_names = ttr_db.keys()

    # Get configurations
    cfg_action, cfg_team_size, cfg_position_name = get_cfg(xls_file_name)

    # Get previous schedule
    prev_schedule = get_prev_schedule(xls_file_name, prev_name, cfg_position_name)
    print_schedule(prev_schedule, cfg_position_name)
    total_new_schedule = [] + prev_schedule # Important: using + to avoid copy by reference

    # Build schedule for N days
    for day in range(DAYS_TO_PLAN):

        # Process previous schedule
        prev_night_list = update_db_with_prev_schedule(valid_names, ttr_db, prev_schedule)

        # Build next day schedule
        new_schedule = build_schedule(prev_schedule, prev_night_list, ttr_db, cfg_action, cfg_team_size)
        print_schedule(new_schedule, cfg_position_name)
        check_for_idle(ttr_db, new_schedule)

        # Get next sheet name
        next_name = get_next_date(prev_name)
        prev_name = next_name

        # Write to XLS file
        if do_write: write_schedule_to_xls(xls_file_name, new_schedule, next_name, cfg_position_name)

        total_new_schedule = total_new_schedule + new_schedule
        prev_schedule      = new_schedule

    # Run checks
    verify        (ttr_db, total_new_schedule)
    check_fairness(ttr_db, total_new_schedule)


##################################################################################
if __name__ == '__main__':
    main()
