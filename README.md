# Webscrapper_for_stocks_nse

Webscrapper using python to abstract stocks details from sites like nseindia.com and bse for daily updated data


Python Dependencies
1.  selenium
2.  pandas
3.  openpyxl
4.  xlsxwriter
5.  BuitifulSoup


Prerequisits
1.  mozilla Firefox ( >>v76.0)
2.  selenium driver (refered geckodriver)
      Download Link https://github.com/mozilla/geckodriver/releases
      Go for latest Release


File Structure

HomeDir
├───company
├───nifty
├───temp
└───turnover

1.  Company-
          ---contains abstracted data of each company listed in Nifty50 
2.  nifty 
          ---contains details of Nify and BankNifty
3.  temp
         ---csv files downloaded form internet to update the data
         ---remove the need to again and again install the same files
4.  turnover
         --- contains turnover of the market day wise
 
