# Stock_Mate
Authors: Alonso Bernal

Module Requirements: yfinance (pip install before running code)

Allows one to generate a simple list of stock recommendations.
Using a GUI to collect user input, and cross-referencing this to the SP500, Stock_Mate can create a personalized portfolio recommendation.
The program accepts .xlsx and .xsl files as input, but only outputs .xsl files.
You will need even a "Simple" .xsl file to work the program. One with the SP500 has been provided.
You only need the first column of this "Simple" file for the program to run.

Similar formats could be utilized to analyze any excel stock list of your choosing.
Be warned however, that a stock list too large a size might prompt an error from the Yahoo API server.
Additionally, the Yahoo API server does not have full information about all stocks currently on the market.
Some trial and error might be required.

The best way to run this program is with a "Full" excel sheet. An example has been provided, and you are encouraged to utilize it.
Creating your own "Full" excel sheet from a "Simple" excel sheet is possible, you simply click "Yes" to the first user prompt.
However, doing so will take an extremely long time due to the nature of the yahoo finance module and the API therein.
For 500 stocks, this can take about 10 minutes.

Therefore, unless the "Full" excel sheet is remarkably outdated, users are encouraged to utilize it when running the program and making minor adjustment thereafter.
Most other directions about format are explained in the GUI.

The output excel file is named: Stock_Mate_Choices.xls

Have fun!

Sincerely,

Your Authors.


*Please note this is a simple recommender, do not make hasty financial decisions.
