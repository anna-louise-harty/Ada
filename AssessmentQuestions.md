#### 1. Find the month in 2015 where the State of Washington had the largest number of storm events. How many days of storm-free weather occurred in that month?  

 ### Answer  
 In Washington the month of December 2015 had the most storm events, with 140 events taking place.  In total there were 8 days of storm free weather in the month of December.

 ### Process  
  - Use Find and Replace in Begining_yearmonth column, replace 2015 with “ “ and sort data by this column
  - Filter to view only the state of Washington
  - Filter to view only January
  - Use Subtotal (2, range) to count the number of storm events in January, repeat for each month (table below)
  - Sort by Beginning day, find days not mentioned in beginning day or end day

| Month | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 |  
 | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |  
| Event | 109 | 19 | 4| 5 | 31 | 25 | 24 | 44 | 17 | 13 | 58 | 140 |  

  Days without storm weather: 14, 25, 26, 27, 28, 29, 30, 31

  ### Notes
  - Could have sorted data using month tab instead of removing 2015.  
  - Could have used filter to record days without storm weather instead of scrolling through data.
  - Repeating for every month was cumbersome.  Could use a macro and expression else/if to count and sum?  

#### 2\. How many storms impacting trees happened between 8PM EST and 8AM EST in 2000?  

 ### Answer  
 There were a total of 2920 storms impacting trees occurring between the hours of 8PM and 8AM EST during the year 2000.  

 ### Process
 - Filter Begin_Time using greater than or less than, realize this does not account for time zones,
 - Run macro to change military times to AM/PM  

          Sub NumberToTime()
            Dim rCell As Range
            Dim iHours As Integer
            Dim iMins As Integer

          For Each rCell In Selection
          If IsNumeric(rCell.Value) And Len(rCell.Value) > 0 Then
              iHours = rCell.Value \ 100
              iMins = rCell.Value Mod 100
              rCell.Value = (iHours + iMins / 60) / 24
              rCell.NumberFormat = "h:mm AM/PM"
          End If
          Next
          End Sub

 - Filter Episode_narrative using contains tree or trees  
 - Filter by time zones (see table)  
 - Filter Begin_Time using Greater than or equal to and less than or equal to (change values depending on time zones for EST & AST use .83333 and .333333, for CST use .7916666 and .2916666 ect)
 - Use =SUBTOTAL (2,range) to count the number of storms

| Zone | EST & AST | CST | HST | MST | PST | SST |
 | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
 | # Events | 1444 | 1232 | 15 | 100 | 122 | 7 |  

 ### Notes  
 - Did not need to run macro, could have used 24 hour clock to get the same results (still learned what a macro was and it is neat).  
 - Could have run a macro using if/else to change time zones and output all times in EST?  Found equation =MOD(time + hours/24),1 but was unsure of how to combine this with if/else in macro.
 - Did not go through the End time data to see if any event began before 8PM but continued after 8 PM, assumed that this was not to be included as it did not fall completely within range.  

#### 3\. In which year (2000 or 2015) did storms have a higher monetary impact within the boundaries of the 13 original colonies?  

 ### Answer  
 In the year 2000 there was a higher monetary impact caused by storms within the boundaries of the 13 original colonies than in the year 2015.  

 ### Process  
 - Filter to see only states that were within the boundaries of the 13 original colonies (Georgia, South Carolina, North Carolina, Virginia, Maryland, Delaware, New Jersey, Pennsylvania, New York, Connecticut, Rhode Island, Massachusetts, New Hampshire, Maine (included in the Massachusetts colony))  
 - Filter Damage to property using contains “k”  
 - In a new column apply formula =SUBSTITUTE(Z1, “K”, “ “)+0  
 - Use =SUBTOTAL (9, range)  
 - Repeat for crops, repeat substituting for "M" (see table)  

| Year | 2000 | 2015 |
| :---: | :---: | :---: |
| Crops K | 16,364,300 | 2,492,000 |
| Crops M | 246,610,000 | 11,000,000 |
| Property K | 58,477,360 | 82,109,700 |
| Property M | 355,390,000 | 189,800,000 |
| Total | 676,841,660 | 285,401,700 |

 ### Notes  
 - Did not account for the part of Maine whose ownership was disputed by the British.
 - Could have used a formula to turn numberK and numberM into 1000 and 1000000 numbers respectively?
 - Could sometimes copy down formula by double clicking cross in the corner, sometimes this did not work.  Need to identify why.
