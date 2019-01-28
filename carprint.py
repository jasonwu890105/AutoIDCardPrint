import pypyodbc
from pyad import aduser
from os import listdir

### Get user lists to check and to print ###
user_tnumber_input_list = input("Please Enter User's tnumber. Please seperate by a comma and a space (eg, t11111, t22222): ")
user_list = user_tnumber_input_list.split(', ')


### Connect to DB, and list all the user Tnumbers ###
conn = pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\t5721793\Desktop\TPG_Card_Data.accdb;')
cursor = conn.cursor()
rows = cursor.execute('select Tnumber from "TPG Card Data"').fetchall()
all_tnumbers = list(map(lambda x: x[0].upper(), rows))


### Define a function if user exists, then change the printercounter to NULL to print ###

#def existing_record(tnumbers):
#    cursor.execute("UPDATE [TPG Card Data] SET printcounter = NULL WHERE Tnumber = ?", tnumbers)
#    cursor.commit()


### Verify if the input users are in the DB Tnumber list or not ###
new_user_list = []

for user in user_list:
    
    if user.upper() in all_tnumbers:
        cursor.execute("UPDATE [TPG Card Data] SET printcounter = NULL WHERE Tnumber = ?", (user,))#Since SQL acccept only inetrable variables, so we need to use the list, not the single user value, refer to https://rogerbinns.github.io/apsw/cursor.html
        cursor.commit() 
        print (f'{user} PrintCounter has been reset to NULL to be printed')


    else:
        new_user_list.append(user)
        try:
            user_AD = aduser.ADUser.from_cn(user)
            user_tnumber_upper = user.upper()
            #user = aduser.ADUser.from_cn(user_tnumber_input_single.upper())

            user_firstname = user_AD.givenname
            user_lastname = user_AD.sn
            user_cardnum = user_tnumber_upper.split('T')[1]
            user_title = user_AD.title
            user_location = user_AD.physicalDeliveryOfficeName

            params = [user_tnumber_upper, user_firstname, user_lastname, user_cardnum, user_title, user_location]
            cursor.execute("insert into [TPG Card Data](Tnumber, FirstName, Surname, CardNumber, Position, Location) values (?, ?, ?, ?, ?, ?)", params)


            cursor.commit()
            print(f'New user {user_AD.displayname} has been Inserted into the Card Print DB')
        except:

            print(f'{user} cannot be found in AD')

print('Checking new users photos.........')

all_photo_in_file_list = listdir(r'\\tpg.local\data\CardExchange\Data\Photos')
all_photo_in_file = list(map(lambda x:x.split('.')[0].upper(), all_photo_in_file_list))
for new_user_photo in new_user_list:
    if new_user_photo in all_photo_in_file:
        print(f'{new_user_photo} has been uploaded')
    else:
        print(fr'Plesae upload {new_user_photo} to "\\tpg.local\data\CardExchange\Data\Photos" folder, and name it as {new_user_photo}')
