'''
Recieves an Excel file as input and outputs (prints) all the emails that is
finds in it.

Example:
> python emails.py contacts.xls > emails.txt
'''
import sys
import occurrences as occ

pattern = r'([a-zA-Z0-9\-\_\+]+\@[a-zA-Z0-9\-\.]+\.[a-zA-Z]+)'


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

    emails = occ.get_occurrences(pattern, data)

    for email in emails:
        print(email)
