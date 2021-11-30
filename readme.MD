Project automates the process of extracting data from itdashboard.gov.

The bot gets a list of agencies and the amount of spending from the main page. Writes the amounts to an excel file and call the sheet "Agencies".

The bot selects one of the agencies, for example, National Science Foundation. Goes to the agency page scrapes a table with all "Individual Investments" and writes it to a new sheet in excel. If the "UII" column contains a link, opens it and download PDF with Business Case (button "Download Business Case PDF").

Store downloaded files and Excel sheet to the root of the output folder
