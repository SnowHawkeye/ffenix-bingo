import argparse
from utils import generate_bingo

parser = argparse.ArgumentParser()
parser.add_argument("--repo", help="name of the items repository")
parser.add_argument("--mask", help="id of the grid mask sheet")
parser.add_argument("--name", help="name of the sheet the grid will be written on")
args = parser.parse_args()

generate_bingo(
    items_repository_sheet_name=args.repo,
    mask_sheet_id=args.mask,
    grid_sheet_name=args.name,
)

# python ffenix_bingo_maker.py --help
# python ffenix_bingo_maker.py --repo "Items" --mask 922974278 --name "Test Grid"

