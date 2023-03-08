# CancelPrintJobsOnErroredPrintersInServer

This script checks for errored printers in a Windows print server and cancels all the jobs related to those printers in some of those states of error:  
  - 1 (Other error)
  - 6 (No toner)
  - 8 (Jammed)
  - 9 (Offline)

This script was created using Powershell baby 
