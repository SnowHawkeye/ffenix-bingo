import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from collections import Counter
import os
import pickle
import json


def load_params():
    file = open("params.json")
    return json.load(file)


def create_service(credentials_path, token_pickle_path):
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = None

    if os.path.exists(token_pickle_path):
        with open(token_pickle_path, 'rb') as token:
            credentials = pickle.load(token)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
            credentials = flow.run_local_server(port=0)
        with open(token_pickle_path, 'wb') as token:
            pickle.dump(credentials, token)

    service = build('sheets', 'v4', credentials=credentials)
    return service


def load_data(service, spreadsheet_id, values_range):
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=spreadsheet_id, range=values_range).execute()
    values_input = result_input.get('values', [])

    if not values_input:
        print('No data found.')

    else:
        return values_input


def load_data_as_dataframe(service, spreadsheet_id, values_range):
    data_raw = load_data(service, spreadsheet_id, values_range)
    data = pd.DataFrame(data_raw[1:], columns=data_raw[0])
    return data


def export_dataframe(service, spreadsheet_id, sheet_name, dataframe):
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='RAW',
        range=sheet_name + '!A1:AA1000',
        body=dict(
            majorDimension='ROWS',
            values=dataframe.T.reset_index().T.values.tolist())
    ).execute()
    print('Sheet successfully Updated')


def duplicate_sheet(service, spreadsheet_id, source_sheet_id, new_sheet_name):
    request_body = {
        "requests": [
            {
                "duplicateSheet": {
                    "sourceSheetId":source_sheet_id,
                    "newSheetName":new_sheet_name,
                    "insertSheetIndex":0
                }
            }
        ]
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=request_body
    ).execute()


def make_item_pool(occurrences, items_repository):
    repo = items_repository.replace("", float("NaN"))
    pool = dict()
    for category in occurrences.keys():
        pool[category] = list(
            repo.dropna(subset=[category])[category] # remove empty values
                .sample(occurrences[category]) # get as many items as needed for the category
        )

    return pool


def make_grid_from_mask(
        service,
        spreadsheet_id,
        items_repository_sheet_name,
        mask_sheet_id,
        grid_sheet_name,
):
    # Load data from API
    duplicate_sheet(service, spreadsheet_id, mask_sheet_id, grid_sheet_name)
    grid_mask = load_data(service, spreadsheet_id, values_range=grid_sheet_name + '!A1:AA1000')
    items_repository = load_data_as_dataframe(service, spreadsheet_id,
                                              values_range=items_repository_sheet_name + '!A1:AA1000')

    # Prepare a pool of items to avoid duplicates
    flat_grid_mask = [cell for row in grid_mask for cell in row]
    category_occurrences = Counter(flat_grid_mask)
    pool = make_item_pool(category_occurrences, items_repository)

    # Create the grid
    grid = grid_mask.copy()
    for i, row in enumerate(grid_mask):
        for j, category in enumerate(row):
            grid[i][j] = pool[category].pop()

    # Update contents in spreadsheet
    request_body = { "values" : grid }

    request = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=grid_sheet_name + '!A1:AA1000',
        valueInputOption='RAW',
        body=request_body)
    request.execute()


def generate_bingo(
        items_repository_sheet_name,
        mask_sheet_id,
        grid_sheet_name,
):
    params = load_params()
    service = create_service(params["credentials_path"], params["token_pickle_path"])
    make_grid_from_mask(service, params["spreadsheet_id"], items_repository_sheet_name, mask_sheet_id, grid_sheet_name)
