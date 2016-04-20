2:19 PM 4/7/2016

Powershell script to test Sage 50 performance during a writeback.
Put your Sage Developer token into the script as marked with "YOUR DEV TOKEN HERE".

1. Run powershell: Windows start prompt / type in: "powershell". Choose Powershell (x86).
	Start as a regular user.

2. Type in: "cd <this script location>".

3. Type in: & "Slow Sage 50 writeback.ps1" and hit Enter.
	Can do the same by just starting typing "slow" and pressing Tab.

4. Green prompts will ask for variables: Server, Company, Invoice number.

== there will be errors during first execution. Open the company and enable access as usually with ShipGear. ==

5. Yellow labels will show the performance metric.

6. Script makes the following changes in the selected invoice:
	Invoice customer note: sets to "Test Customer Note".
	Invoice internal note: sets to "Test Internal Note".
	Invoice freight: sets to 3.14.



There might be an issue with the script: when it runs on a Windows 7 PC it might not find any Sage.Peachtree.API assemblies in GAC.
A workaround is described here:
http://viziblr.com/news/2012/5/16/the-easy-way-to-run-powershell-20-using-net-framework-40.html