# CancelPrintJobsOnErroredPrintersInServer

This script checks for errored printers in a Windows print server and cancels all the jobs related to those printers in some of those states of error:  
  - 1 (Other error)
  - 6 (No toner)
  - 8 (Jammed)
  - 9 (Offline)

This script runs every 5 minutes in a scheduled job as a Windows Task in a Print Server. Whenever it finds printers with some of those errors, it cleans its stored print jobs and notifies the network admins which printers in those states of error had jobs and how many of them were reserved.

This script was created using Powershell baby 
