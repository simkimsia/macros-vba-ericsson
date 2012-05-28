README
======

Normal output file
------------------

To use the macros for normal output,

Unzip the file.

Open BlackBox(Demo).xlsx

In Sheet 1, click Go

It will prompt you for a data input file

Navigate to the input file we've provided and click okay

Give it about 5 mins to run (for 300+ Lines). You can go do something else, except use BlackBox(Demo) or the Input file.

After it is done, it will output an excel file called Output(DD-MMM-YYYY).xls



Damien's Output file
------------------

To use the macros for output that Damien needs,

Unzip the file.

Open BlackBox(Demo).xlsx

In Sheet 1, click Damien

It will prompt you for a data input file

Navigate to the input file we've provided and click okay

Give it about 5 mins to run (for 300+ Lines). You can go do something else, except use BlackBox(Demo) or the Input file.

After it is done, it will output an excel file called 2ndbottom_Output(DD-MMM-YYYY).xls


New Input File (3 Phases)
------------------------

This is another input file that contains 3 phases each eith 1 area, rather than 3 areas in 1 Phase.


Assumptions
For Summary Tab, we've assumed the formulas as follows:
Total Gross = Total Gross of the Tab
Total Net (b4 incentives) = Total Net of the tab
Total Net (after incentives) = Total net (b4 incentives) - Incentives

We expect all areas (level 2 items) to have unique names

For mistakes, we expect mistakes to only happen at the end of an item number, projected downwards i.e
in this series,
[1.1.1
 1.1.1.1
 1.1.1.2
 1.1.3
 1.1.3.1]
the macros will correct it but not the following mistake
[1.1.1
 1.1.1.1
 1.1.2.2 <- unable to correct because there is no parent with this mistake
 1.1.3	 <- able to correct because mistake is at the end
 1.1.3.1]
As long as the mistake happens at the END of an item number and is projected downards, it will be corrected, but not mistakes that occur in the middle of an item number without the parent having the same mistake.

