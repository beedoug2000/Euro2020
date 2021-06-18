import argparse
from openpyxl import load_workbook
import logging

### Set up the logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
s_handler = logging.StreamHandler()
s_formatter = logging.Formatter('%(message)s')
s_handler.setFormatter(s_formatter)
logger.addHandler(s_handler)

### Start of processing
logger.info("Starting the update")

### Get the filename from the command line and open the file for processing
parser = argparse.ArgumentParser(description='Update knock-out stage bracket with team names.', usage='knockout_stage_updater.py <file_name> <name_to_replace> <team_name>')
parser.add_argument('file_name', help='The name of the file to process')
parser.add_argument('name_to_replace', help='The placeholder name in the knock-out bracket')
parser.add_argument('team_name', help='The name of the team to put in the knock-out bracket')
args = parser.parse_args()

filename = args.file_name
name_to_replace = args.name_to_replace
team_name = args.team_name
logger.info("Updating file {}".format(filename))
logger.info("Replacing {} with {}".format(name_to_replace, team_name))

file = None

try:
    file = load_workbook(filename=filename)
except FileNotFoundError:
    raise SystemExit("No such file or directory: '{}'".format(filename))

sheetnames = file.sheetnames
for sheetname in sheetnames:
    if sheetname in ('Leaderboard'):
        logger.debug("'{}' not a knock-out stage sheet - continuing.".format(sheetname))
        continue
    
    logger.info("    Updating {}".format(sheetname))
    for col in file[sheetname].iter_cols(min_row=43, max_row=65, min_col=3, max_col=15):
        for cell in col:
            if cell.value == name_to_replace:
                cell.value = team_name
                break
            
file.save(filename=filename)
