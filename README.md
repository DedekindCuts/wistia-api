# wistia-api
A set of functions to get information about Wistia video views from the Wistia API and write to a Google sheet

## Description

Use these functions to create a button inside a Google Sheet that will allow you to append information from Wistia's API to the spreadsheet.

## Getting Started

### Prerequisites
To get data from Wistia you will need a valid Wistia API key.
Other than that all you need is a blank Google Sheet.

### Usage
First you will need to open the Google Apps Script Editor.
From your blank Google Sheet, open Tools > Script Editor.

![Imgur](https://i.imgur.com/VkMrsU5.png)

Then replace the default code there with our code, and modify it to suit your needs.
**Remember to replace** `YOUR_WISTIA_API_KEY_HERE` **in** `GetWistiaViews` **with your Wistia API key and make sure to set the appropriate time zone in** `GetDates`.
(Make sure to keep the quotation marks.)

![Imgur](https://imgur.com/zcgmf0w.png)

![Imgur](https://imgur.com/NdYRcg9.png)

Now we will add a button to our spreadsheet so that we can run this script whenever we like without having to open the Script Editor.
To do this, just run the function `onOpen`.

![Imgur](https://imgur.com/49Y7NCH.png)

This will add a folder called Wistia to the Menu Bar of our sheet which from which we can run the script.

![Imgur](https://imgur.com/FMkQuAq.png)

Each time you run the script, you will be asked if you only want to retrieve yesterday's data from Wistia.
(The idea being that you will run this script daily to get all the data from the previous day.)
Choosing Yes will then do exactly that.
If you would like to get data from a different time period, you can choose No and another dialog box will ask for the **first** day for which you would like to retrieve data.
For example, if I wanted to retrieve data for the 13th to the 15th of March 2018, I would enter 2018-03-13 as the first date, like so:

![Imgur](https://imgur.com/7MtkTkS.png)

Then another dialog box will ask for the last date, and I would enter 2018-03-15 similarly.
Once the last date is entered, the appropriate data from Wistia will be appended to the bottom of the sheet.
Upon completion, another dialog box will appear confirming that the script has finished.

![Imgur](https://imgur.com/d2zyjzt.png "Yes, I know that the dates are wrong here")

### Customization
First, and most importantly, make sure to replace `YOUR_WISTIA_API_KEY_HERE` with your Wistia API key, or else the script won't work at all.
Also make sure to set the time zone appropriately.

You can also customize the script to include whichever fields that are returned by Wistia's API that you desire.
The full list of available fields can be found [here](https://wistia.com/doc/stats-api#events_list).
If you do wish to change which fields are written to the sheet or how they are written, just alter the definition of `row` in the function `WriteToSheet`.
For example, if I wanted to also include a column showing the country where each event took place, I would add `data[i]["country"]` to `row` as follows:

![Imgur](https://imgur.com/LUvxTAT.png)

It's also important to note that the script as currently written will only write video view events from identified users to the sheet.
To change this behavior, just remove the indicated `if` statement from the function `WriteToSheet` so that the function is as follows:

![Imgur](https://imgur.com/wB57k7A.png)

## Notes

If something doesn't work, check the logs (View > Logs or Cmd + Enter); I've tried to make it so that errors will be logged with a fair level of detail.
Keep in mind that Wistia limits accounts to 100 API calls per minute.
Each page of 100 views counts as 1 call, so this shouldn't be an issue since it takes awhile to even write 1 page of results to the sheet.
Also this was my first experience using Google Apps Script, and [this](https://www.benlcollins.com/apps-script/beginner-apis/) page was extremely useful to me.
I also referenced Wistia's [Stats API documentation](https://wistia.com/doc/stats-api) frequently.
