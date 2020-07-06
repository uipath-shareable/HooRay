<a href="https://www.uipath.com"><img src="https://cdn.icon-icons.com/icons2/21/PNG/256/toons_homer_simpson_homersimpson_2381.png"></a>

# HooRay

> A UiPath robot to help send reminder for public holiday, employee anniversary and employee birthday

## Workflow
<a href="https://github.com/uipath-shareable/HooRay/blob/master/Workflow.jpeg"><img src="https://github.com/uipath-shareable/HooRay/blob/master/Workflow.jpeg"></a>

## Pre-requisite

- UiPath Studio 2019.10.1 or above
  * This is the program platform
- UiPath Orchestrator 2019.10.1 or above
  * This is where we store Sharepoint Credential and Slack Tokens
- Python3
  * Python script is invoked to generate personalized blessing image
  * Please refer to <a href="https://docs.uipath.com/activities/docs/python-scope">link</a> on how to run python script in UiPath
- Slack
  * This is the communication tool where we send reminder message
- Sharepoint
  * This is the repository where we store and download holiday and employee information

### Download

- Clone this repo to your local machine

### Setup

- Update Slack Channel where you send reminder message to
  * in Main.xaml -> Variables -> dict_slack_for_region
- Update Slack admin account who receive error message
  * in Main.xaml -> Variables -> slack_acc_error_report
- Update Sharepoint credential
  * in Download_file_from_sharepoint.xaml -> Variables -> application_id
  * please make sure the format of downloaded public_holiday and employee information files are as expected, you can refer to sample files in downloads/Asean
- Update Slack tokens
  * in Send_slack_msg.xaml & Get_slack_user_ID.xaml -> Variables -> slack_token

## Support

Reach out to me for any question
- wei.wang@uipath.com

---
