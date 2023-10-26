#!/proj/mislcad/areas/DAtools/tools/python/3.10.1/bin/python3

# Status:
# Receive and parse prev_schedule, use for next
# To do:
# - Add check for current schedule for people not served
# - Add sanity check for people at prev_schedule, which are not on the people list
# ---- Add check that TTS is kept
# - Add constraint: if served at prev night (0:00 - 6:00), don't assign tonight
# - Consider getting 2 or more prev schedules
# - Mix couples
# - Add personal constraints

import sys
import os
import random
import pandas as pd
import datetime
import argparse

# For writing XLS file
import openpyxl
from openpyxl.styles               import PatternFill
from openpyxl.utils.units          import cm_to_EMU
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.styles               import Font

# Constants
HOURS_IN_DAY     = 24
TIME_TO_REST     = 6
NUM_OF_POSITIONS = 5
COLUMN_WIDTH     = 27
LINE_WIDTH       = 10 + NUM_OF_POSITIONS*COLUMN_WIDTH

# Generation mode
FIFO_MODE = 1

# Actions:
NONE   = 0
SWAP   = 1
RESIZE = 2

# Colors
PINK   = 0
BLUE   = 1
GREEN  = 2
YELLOW = 3
PURPLE = 4

night_hours = [1, 2, 3, 4]
night_hours_rd = [1, 2, 3, 4]
night_hours_wr = [23, 0, 1, 2, 3, 4]

# Set the random seed (make reproducible)
SEED = 123

##################################################################################
# Utils
##################################################################################

def print_delimiter(): print("#" * LINE_WIDTH)
def error(message):    print('Error: '+message); exit(1)

##################################################################################
# Functions
##################################################################################
def parse_arguments():
    # Create an ArgumentParser object
    parser = argparse.ArgumentParser(description="Command-line parser")

    # Define the command-line arguments
    parser.add_argument("file_name", type=str,            help="XLS file name ")
    parser.add_argument("--seed",    type=int,            help="Seed")
    parser.add_argument("--prev",    type=str,            help="Prev schedule sheet name (optional, default is today's date)")
    parser.add_argument("--next",    type=str,            help="Next schedule sheet name (optional, default is tomorrow's date)")
    parser.add_argument("--write",   action="store_true", help="Do write result to the XLS file")
    parser.add_argument("--fifo",    action="store_true", help="Generate using FIFO mode")
    parser.add_argument("--random",  action="store_true", help="Generate using RANDOM mode")

    # Parse the command-line arguments
    args = parser.parse_args()

    # Access the parsed arguments
    xls_file_name = args.file_name
    prev_name     = args.prev
    next_name     = args.next
    do_write      = args.write
    fifo_mode     = args.fifo
    random_mode   = args.random

    if     fifo_mode and     random_mode:  error("please choose either FIFO mode, or RANDOM mode, can't use both")
    if not fifo_mode and not random_mode:  error("please choose either FIFO mode, or RANDOM mode")
    global FIFO_MODE; FIFO_MODE = 1 if fifo_mode else 0;

    if args.seed: global SEED; SEED = args.seed
    random.seed(SEED)

    if not os.path.exists(xls_file_name):                  error(f"File {xls_file_name} does not exist.")
    if do_write and not os.access(xls_file_name, os.W_OK): error(f"File {xls_file_name} is not writable.")

    return prev_name, next_name, xls_file_name, do_write

##################################################################################
def build_people_db(xls_file_name):
    # Get list of names from XLS
    names = extract_column_from_sheet(xls_file_name, "List of people", "People")

    # Build people DB
    db = []
    for name in names:
        db.append([name[::-1], 0])
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
            # FIXME: sanity (+add hour to error message)
            #if member == 'nan' or member == 'NaN':
            #    error(f"Empty cell in position {position}, hour FIXME")
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

    # Sanity
    #if do_swap and (team_size == 0):
    #   error("For {:02d}:00, swap is required, but team size is not defined (0)".format(hour))

    return do_swap

##################################################################################
# Choose team
def choose_team(hour, night_fifo, day_fifo, night_db, day_db, team_size):
    team = []
    is_night = 1  if hour in night_hours_wr else 0

    #print(f"Choose fifo (is_night = {is_night}), day_fifo = {day_fifo}")
    if FIFO_MODE:
        for i in range(team_size):
            # Choose fifo
            if is_night:
                if len(day_fifo) > 0:
                    fifo = day_fifo
                else:
                    # FIXME: allow override
                    error(f"required to choose team for night watch but both none_fifo and day_fifo are empty")
                    fifo = night_fifo
            else:
                # Is day
                if len(night_fifo) > 0:
                    fifo = night_fifo
                else:
                    fifo = day_fifo

            name = fifo.pop(0)
            team.append(name)
            fifo.append(name)
    else:
        for i in range(team_size):
            # Get list of available people
            if is_night:
                people_available = get_available(day_db)
                if not people_available:
                    people_available = get_available(night_db)
                    if not people_available:
                        error("No available people in neither night_db nor day_db")
            else:
                # Is day
                people_available = get_available(night_db)
                if not people_available:
                    error("For day shift, no people available in night_db. Need to take from day_db (which are already assigned night shift)")
                    people_available = get_available(day_db)
                    if not people_available:
                        error("No available people in neither night_db nor day_db")

            name = random.sample(people_available, 1)[0]
            team.append(name)

            # Update DB before choosing next team member (avoid choosing the same member twice)
            for person in night_db:
                if person[0] == name:
                    person[1] = TIME_TO_REST + 1

            for person in day_db:
                if person[0] == name:
                    person[1] = TIME_TO_REST + 1

    return team

##################################################################################
# Get available people from DB (with TTS == 0)
def get_available(db):
    available = []
    for person in db:
        if person[1] == 0:
            available.append(person[0])
    return available

##################################################################################
# Resize team
def resize_team(hour, night_fifo, day_fifo, night_db, day_db, old_team, new_team_size):

    if new_team_size == 0:
        return ["-"]

    # Create new team list (to avoid modifying the previous hour value, team is passed by reference)
    new_team = old_team.copy()

    old_team_size = len(old_team)
    if old_team_size == new_team_size:
        # FIXME: add error location (position, hour)
        error(f"Resize: old_team_size == new_team_size == {old_team_size}")

    # Resize
    if (new_team_size < old_team_size):
        # Reduce team size
        for i in range(old_team_size-new_team_size):
            random_index = random.randint(0, old_team_size-1)
            released = new_team.pop(random_index)
    else:
        # Increase team size
        for i in range(new_team_size-old_team_size):
            new_member = choose_team(hour, night_fifo, day_fifo, night_db, day_db, 1)
            new_team.append(new_member[0])

    return new_team

##################################################################################
# Make the assignments
def build_schedule(night_fifo, day_fifo, prev_schedule, night_db, day_db, cfg_action, cfg_team_size):
    schedule = [[] for _ in range(HOURS_IN_DAY)]
    teams    = [[] for _ in range(NUM_OF_POSITIONS)]
    team     = []

    for hour in range(HOURS_IN_DAY):
        # Choose teams
        for position in range(NUM_OF_POSITIONS):
            team_size = cfg_team_size[position][hour]
            action    = get_action(cfg_action[position], hour, team_size)
            team      = teams[position]

            if action == SWAP:
                team = choose_team(hour, night_fifo, day_fifo, night_db, day_db, team_size)
            elif action == RESIZE:
                team = resize_team(hour, night_fifo, day_fifo, night_db, day_db, team, team_size)
            elif hour == 0:
                team = prev_schedule[HOURS_IN_DAY-1][position]

            schedule[hour].append(team)
            teams[position] = team

            # Even if there was no swap, the chosen team should get its TTS
            for name in team:
                update_in_db(night_db, name)
                update_in_db(day_db, name)
        #print (f"Night ({hour}): {night_db}")
        #print (f"Day   ({hour}): {day_db}")

        # Update TTS
        update_tts(night_db)
        update_tts(day_db)

    return schedule
##################################################################################
# DB utils
def found_in_db(db, name):
    found = 0
    for person in db:
        if person[0] == name:
            found = 1
    return found

def update_in_db(db, name):
    for person in db:
        if person[0] == name:
            person[1] = TIME_TO_REST+1

def append_to_db(db, name):
    db.append([name, TIME_TO_REST + 1])

def update_tts(db):
    for person in db:
        if person[1] > 0:
            person[1] -= 1


##################################################################################
# Update DB with previous schedule
def update_db_with_prev_schedule(all_db, prev_schedule):
    night_db = []
    day_db   = []
    for hour in range(HOURS_IN_DAY):
        is_night = 1 if hour in night_hours_rd else 0

        for position in range(NUM_OF_POSITIONS):
            team = prev_schedule[hour][position]

            for name in team:
                if is_night:
                    if found_in_db(night_db, name):
                        update_in_db(night_db, name)
                    else:
                        append_to_db(night_db, name)
                        if found_in_db(day_db, name):
                            # Remove from day_db
                            day_db = [item for item in day_db if item[0] != name]

                else:
                    if found_in_db(night_db, name):
                        update_in_db(night_db, name)
                    elif found_in_db(day_db, name):
                        update_in_db(day_db, name)
                    else:
                        append_to_db(day_db, name)

                # Remove name from original db
                all_db = [item for item in all_db if item[0] != name]

        # Update TTS (for each hour, not for each position)
        update_tts(night_db)
        update_tts(day_db)

    # If people remain in all_db (didn't appear in prev_schedule), count them as day_db
    day_db = day_db + all_db

    return night_db, day_db

##################################################################################
# Update FIFO with previous schedule
def update_fifo_with_prev_schedule(people_db, prev_schedule):
    # First should be the people that didn't participate in prev
    # Push all people to FIFO, will be removed if participated
    none_fifo  = []
    night_fifo = []
    day_fifo   = []
    for person in people_db:
        none_fifo.append(person[0])

    # Shuffle
    none_fifo = random.sample(none_fifo, len(none_fifo))

    # Add people that served in prev.
    # For each such person, remove the previous occurrence from the FIFO, insert last
    for hour in range(HOURS_IN_DAY):
        for position in range(NUM_OF_POSITIONS):
           team = prev_schedule[hour][position]
           for name in team:
                if name in none_fifo:
                    # Remove name from FIFO
                    none_fifo = [item for item in none_fifo if item != name]
                    # Append at the end of the relevant fifo
                    if hour in night_hours:
                        if name not in night_fifo:
                            night_fifo.append(name)
                    else:
                        if name not in night_fifo:
                            day_fifo = [item for item in day_fifo if item != name]
                            day_fifo.append(name)

    # People that didn't work yesterday, count as day workers, and are taken before real day workers
    day_fifo = none_fifo + day_fifo

    return night_fifo, day_fifo

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

    # Get sheet name for output (only if not provided by the user
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
# Verify result
def check_fairness(people_db, old_schedule, new_schedule):
    # Collects schedules
    schedule = []
    for hour in range(HOURS_IN_DAY):
        schedule.append(old_schedule[hour] + new_schedule[hour])

    # Init hours_served
    hours_served = {}
    for person in people_db:
        hours_served[person[0]] = 0

    # Calculate hours_served
    for hour in range(HOURS_IN_DAY):
        for team in schedule[hour]:
            for name in team:
                if name != 'nan':
                    hours_served[name] += 1
            hours_served[name] += 1

    # Report
    print_delimiter()
    mode = "FIFO_MODE" if FIFO_MODE else "RANDOM_MODE"
    print(f"Check fairness ({mode})")
    print_delimiter()
    for name in hours_served.keys():
        print(f"Name: {name.ljust(COLUMN_WIDTH)} served: {str(hours_served[name]).ljust(4)}\t"+("*"*hours_served[name]))

    return 1

##################################################################################
# Verify result
def verify(people_db, old_schedule, new_schedule):
    # Collects schedules
    schedule = old_schedule + new_schedule
    print_delimiter()
    print(f"Verify total ({len(schedule)} lines)")

    # Init last_served
    last_served = {}
    for person in people_db:
        last_served[person[0]] = -1

    # Check schedule
    for hour in range(len(schedule)):
        line = schedule[hour]
        for team in line:
            for name in team:
                last_served_hour = last_served[name]
                if last_served_hour != -1:
                    diff = hour - last_served_hour - 1
                    if diff < TIME_TO_REST and diff > 0:
                        error(f"Poor {name} did not get his {TIME_TO_REST} hour rest (served at {last_served_hour}, then at {hour})")
                last_served[name] = hour

##################################################################################
# Main
##################################################################################
def main():

    # Parse script arguments
    prev_name, next_name, xls_file_name, do_write = parse_arguments()

    # Build people DB
    people_db = build_people_db(xls_file_name)

    # Get configurations
    cfg_action, cfg_team_size, cfg_position_name = get_cfg(xls_file_name)

    # Get previous schedule
    prev_schedule = get_prev_schedule(xls_file_name, prev_name, cfg_position_name)
    print_schedule(prev_schedule, cfg_position_name)

    # Process previous schedule
    night_db, day_db     = update_db_with_prev_schedule(people_db, prev_schedule)
    night_fifo, day_fifo = update_fifo_with_prev_schedule(people_db, prev_schedule)

    # Build schedule
    new_schedule = build_schedule(night_fifo, day_fifo, prev_schedule, night_db, day_db, cfg_action, cfg_team_size)
    print_schedule(new_schedule, cfg_position_name)

    # Write to XLS file
    if do_write: write_schedule_to_xls(xls_file_name, new_schedule, next_name, cfg_position_name)

    # Run checks
    verify(people_db, prev_schedule, new_schedule)
    check_fairness(people_db, prev_schedule, new_schedule)


##################################################################################
if __name__ == '__main__':
    main()
