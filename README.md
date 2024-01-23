# Wedding Data Manipulation 
This wedding information has been fabricated for privacy. The csv file is uploaded into Excel for viewing and manipulation.
## Clean, Standardize, and Manipulate.
The data was checked to **Remove Duplicates**. 

The Street and Suite columns were combined into a new Address column by using **Concatenate**.
The City, State Column was separated into different columns using **Split by the comma delimiter**.

A **helper column** was created to summarize Thank You Card statuses by person. The helper column combines First Name, Last Name, and the Thank You Sent columns. 

![Standardized Columns](https://github.com/bvanntruong/EXCEL.Wedding/blob/main/Standardize%20Columns.png)

## Conditional Formatting
For easy viewing, conditional formatting was applied to several columns. 
1. The Side column **text and borders** were formatted Red for Bride and Blue for Groom.
2. A **Data Bar** was added to evaluate the RSVP column.
3. The Donation column has a **Graded Color Scale**, dependent on the amount of donations received.
4. The Thank You Sent column is **filled based on the text value**. "Not Started Yet" is RED, "In Progress" is YELLOW, and "Ready" is GREEN.

![Formatting](https://github.com/bvanntruong/EXCEL.Wedding/blob/main/Conditional%20Formatting.png)

## Additional Formulas & Pivots
The *"Thanks Sent"* worksheet uses VLOOKUP to grab the emails and status from the *"Wedding - Format"* worksheet, displaying it next to the Helper Column. 
It also used the IF operator to evaluate whether the thank you card was sent. If yes, the value will read "DONE". If not, it will display the guest's email address.
There is additional conditional formatting on this worksheet, coloring by status. If it has been completed, it also displays a *strikethrough*.

![Card Statuses](https://github.com/bvanntruong/EXCEL.Wedding/blob/main/thankyoucard.png)

I created Pivot Tables and Charts to visualize our data. I also utilized several IFS formulas to answer exploratory questions.

![Pivots and Formulas](https://github.com/bvanntruong/EXCEL.Wedding/blob/main/Pivots%20and%20IFS%20Q%26A.png)

The *"Weather"* worksheet standardizes the Timestamp column by splitting the Date and Time. The number format for these columns are also standardized. 
Then the Temperatures are converted to Celcius and the Wind Speed is converted to m/s for Canadian guests.
Once these are completed, the data is visualized in charts to forecast the weather for the wedding day.

![Forecast](https://github.com/bvanntruong/EXCEL.Wedding/blob/main/Forecast.png)
