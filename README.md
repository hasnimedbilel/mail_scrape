# Email Scrape

In this war against Covid-19, Cities and whole countries are shut-down in order to flatten the curve of outbreaks of the corona virus. <br>
Everyone of us must, from his place and with his domain of expertise, help the white army to win the battle. <br>
In this context, one of the challenges (with the start of the vaccination compaign) is the management of the huge number of citizens and the processing of their vaccination related information.<br>

## Problem

* Vaccination relative information for each citizen is sent automatically through email to a specific departement in the ministery of health.
* These concerned parties have to extract these information, put in a certain format (Excel table) and further process it.
* The number of emails per dey is huge, and the work is done manually!

## Solution
This script will help scrape all the inbox emails, process them and extract the key elements in the pre-specified format.<br>
* Clone this project
* Open "mail_scrape.py" in your editor
* Change the email credentials to yours:
```
# account credentials
username = "XXX@outlook.com"
password = "XXXXX"
```
* Run the script:<br>
```
> python mail_scrape.py
```
* You'll find the results, within seconds, in the file "output.xlsx"