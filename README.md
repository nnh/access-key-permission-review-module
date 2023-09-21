# Project Title

access-key-permission-review-module

## Project Overview

This project has been created to automate a part of the review process in the "ISF28 入退室管理台帳 自社用".   
Extract data from multiple spreadsheets to identify individuals for review.  

## Project Purpose

The purpose of this project is to streamline the access review process by extracting relevant data, identifying common elements, and creating a user list for review.

## Project Configuration

Before running the project, you need to configure the following script properties:
- `formula`: The formula to retrieve member information.
- `spreadsheetId`: The ID of the "DC Card Key 電子錠カード管理" Google Spreadsheet.
- `targetGroupAddresses`: The email address of the member to be authorized.　Write in the format `["aaa@example.com", "bbb@example.com"]`.
- `formulaForReview`: The formula to be used for access review.

You can set these script properties in the Google Apps Script project associated with your Google Spreadsheet. These properties define essential parameters for the automation process.


## Usage

1. Set up the Google Spreadsheet.
2. Configure script properties.
3. Run the `extractAccessReviewInfo()` function to automate the access review process.

For detailed instructions and usage examples, please refer to "D013-3 SEアシスタントマニュアル（年次）".

