README
======

How to save a Macro Free workbook
---------------------------------
Each input file needs to be saved as a Macro free workbook, so that its macros do not clash with the macros of Blackbox

1. With the input file open, click File -> Save As

2. A small window will pop up. In the "Save as type" field, select Excel Workbook

3. Another window will pop up, saying that the VB project feature can't be saved in a macro-free workbook. Ignore this, and just click Yes to continue saving as a macro-free workbook.

and that's it. You can now run the new macro free input file in Blackbox


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

We assume no orphan nodes. i.e. all nodes have parents except the mother node (apex node)

We expect all areas (level 2 items) to have unique names

For mistakes, we expect no orphan nodes,
i.e in this series,
[1.1.1
 1.1.1.1
 1.1.1.2
 1.1.3
 1.1.3.1]
the macros will correct it but not the following mistake
[1.1.1
 1.1.1.1
 1.1.2.2 <- unable to correct because this is an orphan node
 1.1.3	 <- able to correct because it has a parent
 1.1.3.1]
As long as the no orphan nodes, mistakes will be corrected, but orphan nodes will be ignored and skipped. If there are any blanks in the output file or database, check to input file to see if it is an orphan node.

