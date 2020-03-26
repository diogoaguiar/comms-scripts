'''
Recieves an Excel file as input and outputs (prints) all the mobile phone
numbers that is finds in it.

Example:
> python mobiles.py contacts.xls > mobiles.txt
'''
import sys
import occurrences as occ

pattern = r'(9\d{8}|9\d{2}\s\d{3}\s\d{3})'


if __name__ == '__main__':
    filename = sys.argv[1]

    try:
        data = occ.load_data(filename)
    except occ.InvalidFile as ife:
        print(ife)
        exit(1)
    except occ.InvalidFileType as ifte:
        print(ifte)
        exit(2)

    mobiles = occ.get_occurrences(pattern, data)

    for mobile in mobiles:
        print(mobile)
