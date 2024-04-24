1. For January shifts, run holiday_shifts.py. Else, proceed to 2
	- Assign newbie shifts
	- Replace resigned staff with ‘?’ to randomize predetermined holiday shifts
2. Run shift_sched.py to compute for shift count and generate shift
	- Check gsheet for list of personnel
	- Recompute = TRUE if there's existing shift count for the month (see 3), else FALSE
	- Input previous VPL in previous_vpl
3. Check ShiftCount.xlsx if adjustment is needed then repeat 2. Else, proceed to 4
	- Adjust counts (newbies: 1 MT and CT each)
4. Duplicate current month's sheet in gsheet
	- Reset then edit month name
	- Paste VPL
	- Protect cells: shift summary and VPL
5. Copy MT and CT from MonitoringShift.xlsx to gsheet
