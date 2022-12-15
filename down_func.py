import surveygizmo as sg
import json
import pandas as pd
import math
import os
from datetime import date
import xlsxwriter
import datetime
import openpyxl
import calendar
from threading import Event
import glob
import inspect, os
from tqdm import tqdm
import datetime


## Define Functions

def download_gizmo_without_camp(client, survey_id, page, day_filter_end_day, day_filter_start_day, survey_type):
    if survey_type == "Everything":
        return client.api.surveyresponse.filter(
            'datesubmitted', '<=', day_filter_end_day).filter(
            'datesubmitted', '>=', day_filter_start_day).list(
            survey_id, resultsperpage=500, page=page)
    else:
        return client.api.surveyresponse.filter(
            'datesubmitted', '<=', day_filter_end_day).filter(
            'datesubmitted', '>=', day_filter_start_day).filter(
            'status', '=', survey_type).list(
            survey_id, resultsperpage=500, page=page)


def download_gizmo_with_camp(client, survey_id, page, day_filter_end_day, day_filter_start_day, survey_type,
                             campaign_id):
    if survey_type == "Everything":
        return client.api.surveyresponse.filter(
            'datesubmitted', '<=', day_filter_end_day).filter(
            'datesubmitted', '>=', day_filter_start_day).list(
            survey_id, resultsperpage=500, page=page, campaign_id=campaign_id)
    else:
        return client.api.surveyresponse.filter(
            'datesubmitted', '<=', day_filter_end_day).filter(
            'datesubmitted', '>=', day_filter_start_day).filter(
            'status', '=', survey_type).list(
            survey_id, resultsperpage=500, page=page)


def convert_json_to_xls(input_json):
    df_json = pd.DataFrame()
    df_json['data'] = pd.json_normalize(input_json['data']).T
    df_series = df_json['data']
    return pd.DataFrame(df_series.apply(pd.Series))


def download_contacts(client, survey_id, page, campaign_id):
    return client.api.contact.list(survey_id, resultsperpage=500, page=page, campaign_id=campaign_id)


def flies_to_concat(flies_to_combine):
    comb_df = pd.DataFrame()
    for fle in flies_to_combine:
        fle_df = pd.read_excel(fle)
        comb_df = pd.concat([comb_df, fle_df])
        os.remove(fle)
    return comb_df


def def_time_offset(additional_filters_loc):
    wb_filters = openpyxl.load_workbook(additional_filters_loc, data_only=True)
    sh_filters = wb_filters["Filters"]
    time_offset = sh_filters["A8"].value

    # Define the Additional Filters to find the start/end day of the download
    additional_filters = pd.read_excel(additional_filters_loc)
    day_filter_end_day = additional_filters['End Date'].values.tolist()[0] + pd.DateOffset(hours=time_offset)
    day_filter_start_day = additional_filters['Start Date'].values.tolist()[0] + pd.DateOffset(hours=time_offset)

    return time_offset, day_filter_end_day, day_filter_start_day


def set_start_end_date(additional_filters_loc,
                       day_of_the_week_name):  ## Returns day_filter_end_day day_filter_start_day
    # Define the Additional Filters to find the time_offset

    time_offset = def_time_offset(additional_filters_loc)[0]
    day_filter_end_day = def_time_offset(additional_filters_loc)[1]
    day_filter_start_day = def_time_offset(additional_filters_loc)[2]

    additional_filters = pd.read_excel(additional_filters_loc)
    day_filter_active = additional_filters['Active'].values.tolist()[0]

    today = pd.Timestamp(date.today()) + pd.DateOffset(hours=time_offset)
    yesterday = pd.Timestamp(date.today()) + pd.DateOffset(days=-1) + pd.DateOffset(hours=time_offset)
    thursday_from_sunday = pd.Timestamp(date.today()) + pd.DateOffset(days=-3) + pd.DateOffset(hours=time_offset)

    if day_filter_active == 0:  # Set the main filter
        if day_of_the_week_name == "Sunday":
            return today, thursday_from_sunday
        else:
            return today, yesterday
    else:
        return day_filter_end_day, day_filter_start_day


def get_total_count(responses_page_for_total_count):
    data = [responses_page_for_total_count]
    find_count_start = str(data).find(", 'total_count': ")
    find_count_end = str(data).find(", 'page': ")
    total_count = int(str(data)[find_count_start + 17:find_count_end])
    return math.ceil(total_count / 500)


def gizmo_to_xlsx(input_gizmo, path_file):
    data_page = [input_gizmo]  ##get the data from gizmo

    # make json
    file_name = path_file + '.json'
    file_name = file_name.format('time')
    with open(file_name, 'w', encoding='utf-8') as f:
        json.dump(data_page, f, ensure_ascii=False, indent=4)

    # turning the disgusting json into a beautiful Excel
    input_jason = pd.read_json(path_file + '.json')
    responses_df = convert_json_to_xls(input_jason)
    responses_df.to_excel(path_file + '.xlsx')

    # removing the json file
    os.remove(path_file + '.json')

    # return the responses_df
    return responses_df


def get_campaign_list(responses_no_camp_df):
    campaign_list = responses_no_camp_df['iLinkID'].values.tolist()
    return list(dict.fromkeys(campaign_list))


def find_total_pages_campaign(contacts_first_page):
    start_point = str(contacts_first_page).find(", 'total_pages': ") + 17
    end_point = str(contacts_first_page).find(", 'results_per_page': ")
    return int(str(contacts_first_page)[start_point:end_point])


def add_row_log(log, today_name, main_folder, survey_id, survey_name, campaign_id, file_created, status,
                total_time_downloading,
                day_filter_start_day, day_filter_end_day, complete, partial, deleted, disqualified):
    log_row = pd.DataFrame(
        data={'time_created': [datetime.datetime.now()], 'survey_id': [survey_id], 'survey_name': [survey_name],
              'campaign_id': [campaign_id],
              'file_created': [file_created], 'status': [status], 'total_time_downloading': [total_time_downloading],
              'start_date': [day_filter_start_day],
              'end_date': [day_filter_end_day],
              'complete': [complete], 'partial': [partial], 'deleted': [deleted], 'disqualified': [disqualified],
              'total': [complete + partial + deleted + disqualified]})
    log.to_excel(main_folder + "Log\\log_" + str(today_name) + ".xlsx", index=False)
    return pd.concat([log, log_row])
