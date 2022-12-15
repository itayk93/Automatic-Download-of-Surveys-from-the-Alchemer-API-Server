import surveygizmo as sg
import pandas as pd
import math
from datetime import date
import openpyxl
import calendar
import glob
import traceback
import inspect, os
import datetime
from down_func import download_gizmo_without_camp
from down_func import download_contacts
from down_func import flies_to_concat
from down_func import def_time_offset
from down_func import set_start_end_date
from down_func import get_total_count
from down_func import gizmo_to_xlsx
from down_func import get_campaign_list
from down_func import find_total_pages_campaign
from down_func import add_row_log
from down_func import convert_json_to_xls

# Define Main Folder Location
main_folder = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
folder_name = "API Gizmo Download"
folder_name_loc = main_folder.find(folder_name)
folder_name_len = len(folder_name)
main_folder = main_folder[:folder_name_loc + folder_name_len] + "\\"

# Define the Additional Filters file
additional_filters_loc = main_folder + "Additional Filters.xlsx"

# Define the client login information
wb_filters = openpyxl.load_workbook(additional_filters_loc, data_only=True)
sh_filters = wb_filters["Filters"]
api_version = sh_filters["B13"].value
api_token = sh_filters["B14"].value
api_token_secret = sh_filters["B15"].value

client = sg.SurveyGizmo(
    api_version=api_version,
    api_token=api_token,
    api_token_secret=api_token_secret)

# Define a variable to track the total running time of the script
working_time_script = datetime.datetime.min

# Generate the current date for the log and future download
today_name = str(pd.Timestamp(date.today()))[:-9]

# =========================================================
# Define a filter to select surveys from yesterday only
today = pd.Timestamp(date.today())
yesterday = pd.Timestamp(date.today()) + pd.DateOffset(days=-1)

# Define a filter to select surveys from the weekend only.
thursday_from_sunday = pd.Timestamp(date.today()) + pd.DateOffset(
    days=-3)  # count of the days - supposed to be operated on sunday and show the full weekend

# Create a variable named 'day_of_the_week_name' to define the filtering
day_of_the_week_name = calendar.day_name[date.today().weekday()]
yesterday_name = str(yesterday)[:-9]
# =========================================================

# Define the survey details
survey_list_details = pd.read_excel(main_folder + "SurveyId.xlsx", sheet_name='Surveys')
survey_list = survey_list_details[['Survey ID']].values.tolist()

# Define the log file
log = pd.DataFrame(
    columns=['time_created', 'survey_id', 'survey_name', 'campaign_id', 'file_created', 'status',
             'total_time_downloading', 'start_date', 'end_date', 'complete', 'partial', 'deleted', 'disqualified',
             'total'])

# Download the survey data
for i in range(0, len(survey_list_details)):
    try:
        time_starting = datetime.datetime.now()  # Define the starting time of the download

        # Retrieve the survey details
        df_survey = survey_list_details[survey_list_details.index == i]
        survey_id = df_survey['Survey ID'].values[0]
        survey_name = df_survey['Survey Name'].values[0]
        download_camp_v_x = df_survey['Download Campaign'].values[0]
        survey_type = df_survey['Download Survey Type'].values[0]

        # Specify the days to filter
        day_filter_end_day = set_start_end_date(additional_filters_loc, day_of_the_week_name)[0]
        day_filter_start_day = set_start_end_date(additional_filters_loc, day_of_the_week_name)[1]

        # Set the date to display in the log file
        start_date = day_filter_start_day.strftime("%d/%m/%Y")
        end_date = day_filter_end_day.strftime("%d/%m/%Y")

        # Set an offset of one day on each side to ensure that all surveys are downloaded, accounting for time zone differences
        day_filter_end_day_down = day_filter_end_day + pd.DateOffset(days=+1)
        day_filter_start_day_down = day_filter_start_day + pd.DateOffset(days=-1)

        print("\n" + "Downloading survey " + str(survey_id) + ", Between " + start_date + " - " + end_date)

        # Determine the number of responses to download
        responses_page_for_total_count = download_gizmo_without_camp(client, survey_id, 100000, day_filter_end_day_down,
                                                                     day_filter_start_day_down, survey_type)
        total_pages = get_total_count(responses_page_for_total_count)

        print("total_pages of responses = " + str(total_pages))

        if total_pages == 0:
            log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "", "", "Stopped", "", start_date,
                              end_date, 0, 0, 0, 0)
            print("There are no surveys available today, so we will move on to the next survey")

        else:
            log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "", "", "Started", "", start_date,
                              end_date, "", "", "", "")  # Add a row to the log

            # get the surveys
            for page in range(0, total_pages + 1):
                responses_page = download_gizmo_without_camp(client, survey_id, page, day_filter_end_day_down,
                                                             day_filter_start_day_down, survey_type)
                path_file = main_folder + 'Output\\' + str(survey_id) + '_responses_page_' + str(page)
                responses_no_camp_df = gizmo_to_xlsx(responses_page, path_file)
                log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "",
                                  (str(survey_id) + '_responses_page_' + str(page) + ".xlsx"),
                                  "File Downloaded", "", start_date, end_date, "", "", "", "")  # Add a row to the log
            all_responses = glob.glob(main_folder + 'Output\\' + str(survey_id) + '_responses_page_' + "*.xlsx")

            all_responses_comb = flies_to_concat(all_responses)  # Concatenate all page files into a single file

            try:
                all_responses_comb = all_responses_comb.drop_duplicates(subset=['responseID'])
            except:
                pass

            # Retrieve the time offset that was defined and use it
            time_offset = def_time_offset(additional_filters_loc)[0]

            all_responses_comb['datesubmitted'] = all_responses_comb['datesubmitted'].apply(
                lambda x: (datetime.datetime.strptime(x, '%Y-%m-%d %H:%M:%S') + pd.DateOffset(hours=time_offset)))

            all_responses_comb = all_responses_comb[
                ((all_responses_comb['datesubmitted'] > day_filter_start_day) &
                 (all_responses_comb['datesubmitted'] < day_filter_end_day))
            ]

            # Export the combined responses data
            all_responses_comb.to_excel(main_folder + 'Output\\' + str(survey_id) + '_responses.xlsx', index=False)

            print(str(len(all_responses_comb)) + " surveys were successfully downloaded")

            # Retrieve campaigns that are being used from the responses file
            if download_camp_v_x == "V":
                campaign_list = get_campaign_list(all_responses_comb)
                campaign_list = [item for item in campaign_list if not (math.isnan(item)) == True]

                len_campaign_list = len(campaign_list)
                log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "", "",
                                  "Got Campaign Numbers from the response file", "", start_date, end_date, "", "", "",
                                  "")

                print("The following campaigns are relevant for this survey: " + str(campaign_list))

                # Download the contacts for the campaigns
                for campaign_id in campaign_list:
                    try:
                        campaign_id = int(campaign_id)
                        print("Downloading contacts for campaign " + str(campaign_id))
                        len_sur_for_camp = len(all_responses_comb[all_responses_comb['iLinkID'] == campaign_id])

                        if (len_campaign_list > 5) & (len_sur_for_camp <= 3):
                            message_for_log = "There are only " + str(
                                len_sur_for_camp) + " surveys associated with campaign " + str(
                                campaign_id) + ", and it is not the primary campaign for survey " + str(
                                survey_id) + ", because there are " + str(
                                len_campaign_list) + " other campaigns. Therefore, it will not download its contacts."

                            log = add_row_log(log, today_name, main_folder, survey_id, survey_name, campaign_id, "",
                                              message_for_log, "", start_date, end_date, "", "", "", "")

                            print(message_for_log)

                        else:
                            # Determine the number of pages per contact list in a campaign
                            contacts_first_page = download_contacts(client, survey_id, 1, campaign_id)
                            total_pages = find_total_pages_campaign(contacts_first_page)
                            # print("total_pages of contacts : " + str(total_pages))

                            for page in range(1, total_pages + 1):
                                print("Download page " + str(page) + " for campaign " + str(campaign_id))
                                contacts_page = download_contacts(client, survey_id, page, campaign_id)

                                path_file = main_folder + 'Output\\' + str(survey_id) + '_campaign_' + str(
                                    campaign_id) + '_page_' + str(page)
                                gizmo_to_xlsx(contacts_page, path_file)

                                log = add_row_log(log, today_name, main_folder, survey_id, survey_name, campaign_id,
                                                  str(survey_id) + '_campaign_' + str(campaign_id) + '_page_' + str(
                                                      page) + ".xlsx", "", "", start_date, end_date, "", "", "", "")

                    except:
                        print("Campaign " + str(campaign_id) + " for survey " + str(
                            survey_id) + " is not functioning, so it will be skipped and the next campaign will be processed")

                        log = add_row_log(log, today_name, main_folder, survey_id, survey_name, campaign_id, "",
                                          "The campaign number is not functioning properly", "", start_date, end_date,
                                          "", "", "", "")

                # Combine all contacts into a single data file and save it
                all_contacts = glob.glob(main_folder + 'Output\\' + str(survey_id) + '_campaign_' + "*.xlsx")
                all_contacts_comb = flies_to_concat(all_contacts)
                all_contacts_comb.to_excel(main_folder + 'Output\\' + str(survey_id) + '_contacts.xlsx', index=False)

                print("Finished downloading the contacts")

                # If there are no contacts associated with this campaign, the contacts file will be deleted because it is empty
                if len(all_contacts_comb) == 0:
                    download_camp_v_x = "X"
                    os.remove(main_folder + 'Output\\' + str(survey_id) + '_contacts.xlsx')

                else:
                    pass

            elif download_camp_v_x == "X":
                print("It is not necessary to download contacts from campaigns for this survey")

            # Combine the responses file and the contacts file into a single data file
            if download_camp_v_x == "V":
                final_data = pd.merge(all_responses_comb, all_contacts_comb, left_on='contact_id', right_on='id',
                                      how='outer')
                final_data = final_data.drop_duplicates(subset=['responseID'])

                final_data.to_excel(
                    main_folder + 'Output\\' + str(survey_id) + '_response_contacts_' + yesterday_name + ".xlsx",
                    index=False)

                os.remove(main_folder + 'Output\\' + str(survey_id) + '_responses.xlsx')
                os.remove(main_folder + 'Output\\' + str(survey_id) + '_contacts.xlsx')

            # If there is no campaign (or contacts file) associated with this survey, the data file name will be changed
            elif download_camp_v_x == "X":
                try:
                    os.rename((main_folder + 'Output\\' + str(survey_id) + '_responses.xlsx'),
                              (main_folder + 'Output\\' + str(survey_id) + '_response_' + yesterday_name + ".xlsx"))
                except:
                    os.remove(main_folder + 'Output\\' + str(survey_id) + '_response_' + yesterday_name + ".xlsx")
                    os.rename((main_folder + 'Output\\' + str(survey_id) + '_responses.xlsx'),
                              (main_folder + 'Output\\' + str(survey_id) + '_response_' + yesterday_name + ".xlsx"))

            print("Finished downloading the data for the survey " + str(survey_id))

        # Count the number of downloaded surveys and add it to the log
        if os.path.exists(main_folder + 'Output\\' + str(survey_id) + "_response_contacts_" + yesterday_name + ".xlsx"):
            final_data_count = pd.read_excel(
                main_folder + 'Output\\' + str(survey_id) + "_response_contacts_" + yesterday_name + ".xlsx")
            kind_of_download = "_response_contacts_"

        elif os.path.exists(main_folder + 'Output\\' + str(survey_id) + "_response_" + yesterday_name + ".xlsx"):
            final_data_count = pd.read_excel(
                main_folder + 'Output\\' + str(survey_id) + "_response_" + yesterday_name + ".xlsx")
            kind_of_download = "_response_"

        else:
            final_data_count = pd.DataFrame(columns=['status'])
            kind_of_download = ""

        complete = len(final_data_count[final_data_count['status'] == 'Complete'])
        partial = len(final_data_count[final_data_count['status'] == 'Partial'])
        deleted = len(final_data_count[final_data_count['status'] == 'Deleted'])
        disqualified = len(final_data_count[final_data_count['status'] == 'Disqualified'])

        log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "",
                          str(survey_id) + kind_of_download + yesterday_name + ".xlsx",
                          "Finished Downloading", "", start_date, end_date, complete, partial, deleted, disqualified)

        time_ending = datetime.datetime.now()
        total_time_downloading = time_ending - time_starting

        print("\nSurvey " + str(survey_id) + " Summery:")
        print(str(complete) + " Completed Surveys")
        print(str(partial) + " Partial Surveys")
        print(str(deleted) + " Deleted Surveys")
        print(str(disqualified) + " Disqualified Surveys")
        print(str(complete + partial + deleted + disqualified) + " Total Surveys")

        # Calculate the total time it took for the script to download the survey and record the result in the log
        log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "", "",
                          "Total time Downloading",
                          total_time_downloading, start_date, end_date, "", "", "", "")
        print("0" + str(total_time_downloading * 24)[:-7] + " Minutes To Download The Surveys")

        # Add the time it took for the script to download the survey to the calculation of the total time of download that will be shown at the end.
        working_time_script = working_time_script + total_time_downloading*24
    except:
        sur_crash_log_location = main_folder + "Survey Crash Log\\"
        current_time_log = datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S").replace("-", "").replace(":", "").replace(" ", "_")
        with open(sur_crash_log_location + str(survey_id) + "_" + str(current_time_log) + ".txt", "w") as logfile:
            logfile.write("Survey id : " + str(survey_id) + "\n" +
                          "Crashed at " + str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")) + "\n\n")
            traceback.print_exc(file=logfile)

        log = add_row_log(log, today_name, main_folder, survey_id, survey_name, "",
                          sur_crash_log_location + str(survey_id) + "_" + str(current_time_log) + ".txt",
                          "Download Crashed", "", start_date, end_date, "", "", "", "")

        print("Couldn't Download Survey " + str(survey_id))

#Display the total time it took to download the surveys and record the result in the log file
print("\nThe total time it took to download all of the surveys is " + str(working_time_script)[11:-7])

log = add_row_log(log, today_name, main_folder, "", "", "", "",
                  "The total time it took to download all of the surveys",
                  working_time_script, start_date, end_date, "", "", "", "")
