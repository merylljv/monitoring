# Monitoring shift generator

## Requirements
 - Access to monitoring shift google sheet
 - Python 3.12.x

## Usage
1. If generating for the start of the year (January shift), run ```holiday_shifts.py```. Else, proceed to step **2**
	- Assign newbie shifts (include x number of newbies as placeholder in personnel sheet)
	- Replace resigned staff with ‘?’ to randomize predetermined holiday shifts
2. Remove ```newbies``` in personnel list. Run ```shift_sched.py``` to compute for shift count and generate shift assignments
	- ```Recompute``` - ```TRUE``` if there's an existing shift count for the month (see step **3**), ```else FALSE```
	- Input previous VPL in ```previous_vpl```
3. Check **ShiftCount.xlsx** if adjustment is needed then repeat step **2**. Else, proceed to step **4**
	- Adjust counts (newbies: 1 MT and CT each)
4. Duplicate current month's sheet in gsheet
	- Reset then edit month name
	- Paste VPL
	- Protect cells: shift summary and VPL
5. Copy MT and CT from MonitoringShift.xlsx to gsheet
