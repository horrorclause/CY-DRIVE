#!/usr/bin/env python3
'''
DestiZerg or Cy-Drive
'''

from datetime import date
import sqlite3
from openpyxl import Workbook
from string import ascii_uppercase as alpha
from time import sleep

version = "1.0.0"
# Database and connections to it
database = 'test1.db'

conn = sqlite3.connect(database)
c = conn.cursor()

print("""
+--------------------------------------+
|  Connected to database successfully  |
+--------------------------------------+
""")

# Dictionary of Company list and distance in miles. Reference point is Cyzerg office
company_names_and_miles = {
    'alexander montessori': 16,
    'asia shipping': 3.1,
    'codotrans': 3.6,
    'tms': 19,
    'wtdc': 4.3,
    'brandon brokerage': 11,
    'south fl ac': 3.1,
    'cosmedical': 38,
    'interport': 3.4,
    'terra global': 17,
    'gloval': 2.2,
    'bringer': 3.7,
    'gava': 4,
    'elle logistics': 4.2,
    'lac': 3.5,
    'mag wholesale': 8.2,
    'gap': 5.6,
    'interworld': 5,
    'dmr': .7,
    'designs by nature': 19.2
}

""" 
CREATE TABLE - TEST 
c.execute(('''CREATE TABLE test(   
            id       integer PRIMARY KEY,
            customer   text     NOT NULL,
            miles      integer  NOT NULL,
            date       date     NOT NULL
            )'''))
           
conn.commit() 
ADD COLUMN

c.execute('''ALTER TABLE test ADD COLUMN reimbursement INTEGER''')
"""


def destination():

    q = True

    while q:

        choice = input('\nWhere are you going?\n>> ').lower()

        # Checks to see if choice is in company dict.
        if choice in company_names_and_miles:
            print('\n{} is {} miles away'.format(choice.title(), str(company_names_and_miles[choice])))

            # Logic for whether to save record or not
            n = True

            while n:
                ask = input("\nDo you want to save this record? Y/N\n>> ").lower()
                if ask == 'y' or ask == 'yes':
                    add_records(choice)   # Adds destination as new entry into DB
                    n = False
                elif ask == "n" or ask == "no":    # If you select No the application quits.
                    print("\nOK")
                    n = False

                else:
                    print('\nPlease choose either "Y" or "N"\n')    # loops back so a choice can be made
            q = False

        else:
            print('\nThat company is not listed')


# Adds records to db
def add_records(customer, date_info=date.today()):

    miles = company_names_and_miles[customer]
    cash = round((miles * .57), 2)
    c.execute("INSERT INTO test (customer, miles, date, reimbursement) \
               VALUES (?,?,?,?)", (customer.title(), miles, date_info, cash))

    conn.commit()

    print("\n#------------> Record added successfully <------------#")


# Get information from DB
def get_record():

    choice = input("""
        Please make a selection:
    
            1) All Records
            2) Specific Date
            3) Range
            4) Exit
    
>> """)

    # All records
    if choice == '1':
        count = 0
        info = c.execute('SELECT id, customer, miles, date, reimbursement FROM TEST')

        print("=" * 20)
        print("Retrieving records")
        print("=" * 20, '\n')

        for row in info:
            print("ID= ", row[0])
            print("CUSTOMER= ", row[1])
            print("MILES= ", row[2])
            print("DATE= ", row[3])
            print("REIMBURS.= $", row[4])
            print('--------------------\n')
            count += 1

        print('\n')
        print('*@' * 21, "\n* All {} records retrieved successfully  *".format(count))
        print("*@" * 21)

    # Specific Date
    elif choice == '2':
        date_choice = input("Enter date (ex. YYYY-MM-DD) \n>> ")
        grab_date = c.execute('SELECT CUSTOMER,MILES,DATE,REIMBURSEMENT from test WHERE DATE = ?', (date_choice,))
        for x in grab_date:
            print("---------------")
            print("{} {} miles, {}, ${}".format(x[0], x[1], x[2], x[3]))

    # Range of dates
    elif choice == '3':
        beginning_date = input("Beginning Date (YYYY-MM-DD) :\n>> ").title()
        end_date = input("End Date (YYYY-MM-DD) :\n>> ").title()

        date_range = c.execute('SELECT DATE,CUSTOMER,MILES,REIMBURSEMENT from test WHERE DATE BETWEEN ? AND ?', (beginning_date, end_date))
        count = 0
        for r in date_range:
            record = "{} :: {} miles :: {} :: ${}".format(r[0], r[1], r[2], r[3])
            print(record)
            print("-"*len(record))
            count += 1

        print("\n[+] {} Total Records Retrieved between {} and {}".format(count, beginning_date, end_date))

    # TODO Loop back to beginning if they make a typo
    else:
        print("Wrong choice")

        #exit()


# Creation of excel spreadsheet for exporting
def create_workbook():

    try:

        cursor_work = conn.cursor()

        workbook = Workbook()
        sheet = workbook.active
        # Header of Excel File
        sheet["A1"] = "ID"
        sheet["B1"] = "CUSTOMER"
        sheet["C1"] = "MILES"
        sheet["D1"] = "DATE"
        sheet["E1"] = "REIMBURSEMENT"

        workbook_name = input("Save workbook as >>  ")  # User can customize workbook name
        count = 1
        print("[+] Building Excel SpreadSheet\n")
        # Cycles through excel cells inserting DB information
        for row in cursor_work.execute('SELECT * FROM test'):
            for i in range(5):
                sheet[str(alpha[i])+str(count+1)] = row[i]

            count += 1

            workbook.save(filename=workbook_name + '.xlsx')    # Exported excel file

            # Prints ID# Customer Miles Date Reimbursement
            print("[+] {}  {}  {}  {}  {} ".format(row[0], row[1], row[2], row[3], row[4]))
            print("    Copied over successfully")

        print("\n+----- Export completed -----+")
        cursor_work.close()

    except PermissionError:
        print("[-] The file is currently open, please close it and try again.")


print("""\n
<+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+>



 ██████╗██╗   ██╗          ██████╗ ██████╗ ██╗██╗   ██╗███████╗                
██╔════╝╚██╗ ██╔╝          ██╔══██╗██╔══██╗██║██║   ██║██╔════╝                
██║      ╚████╔╝   ██████  ██║  ██║██████╔╝██║██║   ██║█████╗                  
██║       ╚██╔╝            ██║  ██║██╔══██╗██║╚██╗ ██╔╝██╔══╝                  
╚██████╗   ██║             ██████╔╝██║  ██║██║ ╚████╔╝ ███████╗                
 ╚═════╝   ╚═╝             ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═══╝  ╚══════╝                

                                                        {}

<+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+><+>                                
""".format(version))  # Title Cy-Drive

# Logic for app
p = True
# TODO go through each selection and ensure the choice has an option to go back to main menu
while p:
    print("""
    Select what you would like to do:
    
            1) Add a destination
            2) View records
            3) Export to Excel
            4) Exit
    """)    # Options for app
    choice = str(input(">> "))   # User makes their choice

    if choice == '1':
        destination()
        #   Loop if user wants to do other things after adding destination
        z = True
        while z:
            more = str(input("\nIs there anything else you'd like to do? (Y/N)  "))

            if more.lower() == 'y' or more.lower() == 'yes':    # Continue back to main menu
                print("OK")
                z = False    # Breaks z loop to loop back to main app

            elif more.lower() == 'n' or more.lower() == 'no':   # User is done, exits app
                # Close connection to db
                conn.close()
                print("\n[+] Leaving already?")
                print("""|
|
+-- { Database has been closed } --+
                                   |
                                   |
                           [  BYE  ]""")
                exit()  # Exits code

            else:
                print("\nPlease select either Y or N")  # User made a wrong entry

        #print('---- end of z loop ----')

    elif choice == '2':
        get_record()
        zx = True
        while zx:
            more = str(input("\nIs there anything else you'd like to do? (Y/N)  "))

            if more.lower() == 'y' or more.lower() == 'yes':  # Continue back to main menu
                print("OK")
                zx = False  # Breaks z loop to loop back to main app

            elif more.lower() == 'n' or more.lower() == 'no':  # User is done, exits app
                # Close connection to db
                conn.close()
                print("\n[+] Leaving already?")
                print("""|
|
+-- { Database has been closed } --+\n""")
                for x in "GOODBYE":
                    print(x, end=' ')
                    sleep(.25)
                exit()  # Exits code

            else:
                print("\nPlease select either Y or N")  # User made a wrong entry

        print('---- end of zx loop ----')
        #p = False  # Kills p loop

    elif choice == '3':
        create_workbook()
        p = False
    elif choice == '4':
        p = False
    else:
        print('Please select a correct option.')



#destination()
#add_records()
#get_record()
#create_workbook()

# Close connection to db
conn.close()
print("\n[+] Leaving already?")
print("""|
|
+-- { Database has been closed } --+\n""")
for x in "GOODBYE":
    print(x, end=' ')
    sleep(.25)

exit()
