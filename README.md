In the script that I wrote, surveys can be automatically downloaded from the API server of Alchemer (formerly known as Survey Gizmo).

Instructions:
1. Edit the SurveyId file to include the desired survey IDs.
2. Specify additional filters, such as download dates and client information for the API, in the Additional Filters file.
3. Run the script from the file Download Surveys.py.

What the script does:
1. The script will download the surveys from the specified date range, page by page, with each page containing up to 500 rows, and concatenate them into a single data file of responders to the survey.
2. The script will then extract the campaign IDs from the survey data, using the specified column (by default, this is the "iLinkID" column). If a survey ID is specified in the SurveyId file not to search for campaign IDs, the script will skip to the next survey.
3. The script will create a list of the campaign IDs and download the contacts for each campaign, page by page, with each page containing up to 500 rows. The downloaded contacts will be in JSON format and will be converted to XLSX.
4. If a campaign ID is not found or if there are more than three campaigns in the survey data, the script will skip to the next campaign to save time.
5. After downloading all the surveys and contacts, the script will merge the files into a single full data file.
6. If a download crashes, a new log will be created in the Survey Crash Log folder.

Note: By default, this script is designed to download surveys from the previous day, or the previous Thursday to Saturday if it is the weekend (based on Israeli time). You can modify the dates in the Additional Filters file to download surveys from any desired date range.

To modify the SurveyId file:
1. In the "Survey ID" column, write the number of the survey. This can be found in the URL line when viewing the survey in the Alchemer website.
2. In the "Survey Name" column, write the name of the survey for your personal use.
3. In the "Download Campaign" column, write either "V" or "X" depending on whether you want to download the contact list for the associated campaign.
4. In the "Download Survey Type" column, specify which type of surveys you want to download. If you write "Everything", the script will download all types of surveys (complete, partial, deleted, disqualified). If you want to download a specific type of survey, specify it here (e.g. "partial" to download only partial surveys).

To modify the Additional Filters file:
1. If you want to set custom dates for the script to download surveys, write "1" under the "Active" column (cell A2). If you leave this as "0", the script will download surveys using the default dates.
2. To set the Time Offset, which determines the date that the surveys will be shown as, change the number in the corresponding cell. This is useful if you want the date to match your local time zone, as opposed to the time zone used by Alchemer.
3. To make the script work, you must specify your "api_token" and "api_token_secret" in the corresponding cells (B14 and B15). These can be found in your user settings in Alchemer.
4. You can also change the "api_version" if desired, but it is recommended to use v4 rather than v5.
