Readme written by Mark, March of 2021.

This application was written in 2005 for a company named Seismic Micro-Technology.  They had a room full of machines running tests against the "Kingdom" application suite.  I do not think they care that I've posted this to my GIT since Microsoft Visual Test automation tool is no longer supported (it wasn't even supported whey I wrote this utility!), and so the test suite has surely been replaced or retired after so long.

This application's sole purpose was to automate the execution of the test suite from a remote system.  The user could select the machines they wanted to run tests on, and pass/fail data was displayed when finished (as I recall), and all the log files were collected into one location for parsing to debug problems. 

The primary driver was "runTests.hta", which was a "HyperText Application".  It has a gui frontend written in HTML, and VBScript to handle user interaction and test kickoff.  I do not remember details about the smaller files that are part of this app, but I remember the need to kill rogue processes on remote machines in order so the test pass could complete properly.
