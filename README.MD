Readme written by Mark, March of 2021.

NOTE: Due to Microsoft security measures, this program ... nay, VBScript as a whole, no longer launches on modern Windows systems, as far as I know.  VBScript has been obsoleted.  Nevertheless, I wrote this app so here it is.

It was written in 2005 for a company I previously worked for under contract.  They had a room full of machines running tests against their flagship seismic data collection software suite.  They don't care that I've posted this since the Microsoft Visual Test automation software is no longer supported (it wasn't even supported when they asked me to write this frontend utility), and so the test suite has surely been replaced or retired after so long.

This app's sole purpose was to automate the execution and results collection for a test suite - using one remote system.  The user could select the machine(s) they wanted to run tests on (usually all of them), and the pass/fail log files with desriptions were moved to the remote machine when on completion.  Logs were previously stored separately on each of the 25-30 machines. Reviewing the issues list was obviously much easier this way.  It's like if your yard crew had consecutive customers and could mow all your neighbors' grass as well your own - 15-30 homes in a row.  Look at how much faster they could earn their money! 

The primary driver was "runTests.hta", which was a HyperText Application, aka an HTA app.  It has a GUI frontend written in HTML, and VBScript code to handle user interaction and test kickoff.  I do not remember details about the smaller files that are part of this app, but I remember the need to kill rogue processes on problem machines so the test pass could complete properly on the rest of the test systems.  
