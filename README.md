Send XLS Results



if anything fails  this will show the debug output:

index=_internal sourcetype=sendxlsresults:log

Testing:
| sendalert sendxlsresults_alert param.recipient="" param.xlsfilename="" param.subject="" param.sender=""

#####
13.2.2020: Update xlwt to 1.3
