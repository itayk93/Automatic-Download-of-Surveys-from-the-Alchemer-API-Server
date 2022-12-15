Automatic Download of Surveys from the Alchemer API Server:
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
