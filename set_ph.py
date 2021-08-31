from datetime import date
from json import dump

print('Set public holidays for a specific year (for use with manhour.py).')
year = int(input('Year: '))  # add verification
dates = []

while True:
    entry = input('Date (DD/MM): ')
    if entry.strip() == '':
        break
    day = int(entry[:2])  # change to get values on either side of /
    month = int(entry[-2:])
    try:
        new_date = date(year, month, day)
    except ValueError:
        print('Invalid entry. Try again.')
        continue
    confirm = input('Confirm (y/n)? ')
    if confirm.strip() == 'y':
        if new_date in dates:
            print('Date already exists.')  # move before confirm
            continue
        else:
            dates.append(new_date)
            print('Date added.')
            continue
    else:
        print('Entry cancelled.')
        continue

with open(f'ph_{year}.json', 'w') as file:  # overwrites any existing file
    dates_str =[]
    for date in dates:
        dates_str.append(f'{date.year}/{date.month}/{date.day}')
    dump(dates_str, file)
    print(file)
print('Saved and exiting.')
