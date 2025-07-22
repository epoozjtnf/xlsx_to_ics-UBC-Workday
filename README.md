# xlsx_to_ics-UBC-Workday
A dependency free python script that converts xlsx file for registered courses from workday to a ics calendar.

<img width="749" height="353" alt="Screenshot 2025-07-19 at 5 53 45 PM" src="https://github.com/user-attachments/assets/908d97e8-f3d2-4f6e-bbfc-75f11616a84f" />

## Requirement
Python on the system.
## How to use
Put xlsx file from Workday and the xlsx_to_ics.py in the same folder (if you just downloaded them, they are probably already both inside the Downloads folder)
Open up terminal and enter:
```
cd ~/Downloads
```
Or any other folder as long as the two files are in them and you are running the following line inside that folder:
```
pip install -r requirements.txt
python3 xlsx_to_ics.py
```

If you encounter an error installing packages, run the following line:
```
python3 xlsx_to_ics.py --tz America/Vancouver  # Replace 'America/Vancouver' with your own timezone, e.g., 'Europe/London' or 'Asia/Tokyo'```