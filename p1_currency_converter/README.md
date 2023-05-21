## Problem Statement

Create a user form that will enable one to use real-time exchange rates to convert currency from one unit to another. Automate a data query in Excel, and link the results to an easy-to-use VBA user form. User should be able to enter in any historic date (within reason) to obtain the exchange rates/converted values for times in the past and to generate a plot of exchange rates over the past 30 days.

## Requirements

1. You can start with a list of all the currencies and their descriptions (e.g., “USD” in column A and “US Dollar” in column B, etc.) on a visible worksheet, but the user should never see the imported data.  Start with a hidden sheet, which is then unhidden, data imported to it, data used, but then the sheet is hidden.  Use Application.ScreenUpdating = False such that the user never sees anything that’s going on “behind the scenes”.

2. By default, the date shown should be the current date.

3. Your currency converter user form should successfully make a variety of different exchanges, and display the result in a message box.  The grader should verify that several of these conversions work (choose randomly and compare to current data on http://www.xe.com/currencytables/).

4. Your user form should work on different dates.

5. The Quit button should work perfectly.  When the form is re-opened, you should check to make sure that the combo boxes do not have repeat values.

6. Your user form should successfully plot the last 30 days of exchange rates. 

7. Input validation:
   - If the user inputs a negative or zero currency, it should provide an error.
   - If the user leaves the input field blank and tries to run the button, it should display an error.
   - If the user puts text in the input field, it will display an error.
   - If the date field is not a date, then an error should be displayed