# OutlookTimeReporting
A VBA macro to create a time report based on tagged calendar events in outlook.


# How to
1. Make sure to have events in the calender with categories
1. Run macro and add start and end date when prompted
1. See result (planned is coded to 8 hours per day every weekday)

![Alt text](/images/QuickReport.PNG?raw=true "Report form clipboard.")

1. Note that the result is also added to clipboard and can be added to excel.

![Alt text](/images/DetailedReport.PNG?raw=true "Report form clipboard.")

Tip: A custom button can be added to the ribbon i outlook to trigger the macro.

# ToDo
- Use category as a simple category and add the meeting title to another column in the report.
- Fix hard coded swedish weekday value to a non-language dependant format.
