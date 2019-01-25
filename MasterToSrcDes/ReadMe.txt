Executable: Depotwise-MasterToSrcDes_v2.0
Version: 2.0

Input file can be selected in the browse window.
Provide correct depot name(it is case sensitive) in the command prompt when asked.  (this name will be used in the output filename.)

Column names that will be considered from the input(anyother column will be ignored. name the columns as below):
1. English Stop Name
2. Marathi Stop Name


Output excel will be in the working directory of Executable: MasterToSrcDes_v1.2
Sample Outputfile: "Wadala output.xlsx". This excel will have two sheets,
								  sheet1: "FromToData": Processed data
								  sheet2: "Ord Data"  : Full data related to specified depot.

Note: If the input is wrong, either the excel path, or the depot specified does not exist in the excel, the program will exit automatically. Otherwise output excel will be created.