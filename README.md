Maelstrom Timesheet Python Updater
===================

Update a Google Spreadsheet (Maelstrom-Research Timesheet format) from with python script

Installation
===

- [Install GSpread](https://github.com/burnash/gspread)


Usage
===

	python pytime.py [PROJECT] [--TASK=NB_HOURS] [--OPTION]

Projects: maelstrom, bioshare, p3g, clsa, cpac, bbmri, ialsa, cihr, interconnect, mrc

Tasks: coordination, dev_other, datashield, opal, onyx, mica

Options: day, month

Examples:
===

	# Enter time for current day
	python pytime.py maelstrom --mica 3 --coordination 2.5 --opal 2

	# Enter time for yesterday
	python pytime.py maelstrom --onyx 3.5 --opal 4 --day -1

	# Enter time for a specific day: August 5th
	python pytime.py bioshare --datashield 3 --dev_other 4.5 --day 5 --month 8