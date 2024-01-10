import sys
import win32com.client as w32
from os import system

filters_list = ['subject', 'sender', 'sender and subject']


## FUNCTIONS
def get_user_selection():
    '''gets the operation the user wants to perform and returns it.'''
    
    system('cls')    
    print("Insert the number of the function you want to do based on the options below: ")
    print("1 - Search and list mail\n2 - Create a new mail folder\n3 - Move mail\n4 - Mark mail as read\n5 - Delete mail\n6 - Mark ALL as read\n7 - Delete ALL mail\n8 - End program")
    menu_selection = input("")

    return menu_selection


def get_account():
    '''gets and returns the user account. must be str with a valid mail @outlook/@live/@hotmail.'''
    
    print("Welcome to the mail organizer. ")
    user_account = input("Insert your outlook/live/hotmail account: ").lower()  
    return user_account


def enter_application(account):
    '''takes the account from the user input to create the >inbox< object which contains all mail and returns it.
    if the account is not logged on the outlook app, it will raise an exception and end the program.'''
    
    try:
        outlook = w32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        account = outlook.Folders[account]
        inbox = account.Folders['Caixa de Entrada']

    except Exception:
        system('cls')
        print(f"Error: the account you inserted is not available on the Windows Outlook app.\nPlease make sure your account is logged and the Outlook app is up to date.")
        sys.exit()
    
    return inbox


def get_new_folder_name():
    '''gets the user input to create a new folder'''
    
    return input("Insert the name of the folder you want to create: ")


def create_folder(inbox, new_folder_name):
    '''takes the account from the user input to retrieve all the mail folders'''

    try:
        new_folder = inbox.Folders.Add(new_folder_name)
        print(f"Folder '{new_folder_name}' created successfully.")
        return new_folder
    
    except Exception as e:
        print(f"An error occurred: {e}")
        return None


def get_folder_name():
    '''gets and returns the user input to select the folder he wants to use on the selected operation'''

    return input("Insert the name of the folder you want to use: ")


def move_to_folder(folder_name, mail_list, inbox):
    '''takes name of the folder the user wants to move his mail to and the list of retrieved mail from the previous retrieval method and moves it'''
    try:
        for mail in mail_list:
            mail.Move(inbox.Folders(folder_name))
    
    except Exception as e:
        print(f"An error has occurred while trying to move your mail. Error: {e}")
    

def mark_mail_as_read(mail_list):
    pass


def get_all_mail(inbox):
    '''retrieves all mail from the inbox and puts into a >mails< object'''

    mails = inbox.Items.Restrict("[ReceivedTime] >= '01/01/2007'")
    # items = list(items)[:200]

    return mails


def select_type_mail(filters_list):
    '''takes the list of supported filters from the filters_list and let the user chooses how he wants to filter his inbox.'''

    print("Insert the type of mail you want to filter by the number from the list below: ")
    counter = 0

    for filter in filters_list:
        counter += 1        
        print(f"{counter} - {filter}")

    user_input = int(input(""))

    if user_input <= len(filters_list):
        return filters_list[user_input - 1]
    
    else:
        print(f"Type invalid or not supported yet. Supported filter types: {filters_list}")    
        return None


def get_sender_mail(items):
    '''gets the items object to retrieve all the inbox's mail sent from the sender the user inputs and and prints them. then return the list of emails.''' 

    system('cls')
    mail_count = 0
    user_input = []
    sender = ""
    emails = []

    while sender != "end":
        sender = input("Insert the sender's email address you want to filter from your inbox (insert 'end' to finish): ").lower()
        
        if sender != "end":
            user_input.append(sender)

    for address in user_input:        
        emails.extend([
            {
                'Subject': m.Subject,
                'SenderEmailAddress': m.SenderEmailAddress,
                'ReceivedDate': m.ReceivedTime.date()
            }
            for m in items if address in m.SenderEmailAddress.lower()
        ])
    
    system('cls')

    for email in emails:
        print(f"Email: {email['Subject']} || from: {email['SenderEmailAddress']} || date: {email['ReceivedDate']}")
        mail_count += 1
    
    print(f"\n{mail_count} mail were found using the {user_input}'s search term")
    return emails


def get_mail_by_subject(items):
    '''gets the items object to retrieve all the inbox's mail with the user inputs on the subject and and prints them. then return the list of emails.''' 
    
    system('cls')
    mail_count = 0
    user_input = []    
    term = ""
    emails = []

    while term != "end":
        term = input("Insert the term on the mail's subject you want to filter from your inbox (insert 'end' to finish): ").lower()
        
        if term != "end":
            user_input.append(term)
    
    for search_term in user_input:
        emails.extend([
            {
            'Subject': m.Subject,
            'SenderEmailAddress': m.SenderEmailAddress,
            'ReceivedDate': m.ReceivedTime.date()
        }
        for m in items if search_term in m.Subject.lower()
    ])

    system('cls')

    for email in emails:
        print(f"Email: {email['Subject']} || from: {email['SenderEmailAddress']} || date: {email['ReceivedDate']}")
        mail_count += 1

    print(f"\n{mail_count} mail were found using the {user_input}'s search term")
    return emails


def get_mail_sender_and_subject(items):
    '''gets the items object to retrieve all the inbox's mail from an specific sender with an specific subject from the user input'''

    system('cls')
    mail_count = 0
    sender = input("Insert the sender adress you want to search: ")
    subject = input("Insert the content on the mail's subject you want to search from the sender: ")
    emails = [m for m in items if sender.lower() in m.SenderEmailAddress.lower() and subject.lower() in m.Subject.lower()]

    for email in emails:
        print(f"Email: {email} || from: {email.SenderEmailAddress} || date: {email.ReceivedTime.date()}")
        mail_count += 1

    print(f"{mail_count} mail were found combining the {sender} and {subject} on the search")
    return emails


def check_result(result, mails):
    '''gets the selected filter from the selection and returns the list of mail from this search'''

    if result is not None:
        if result == filters_list[0]:
            filtered_mail = get_mail_by_subject(mails)
            return filtered_mail
        elif result == filters_list[1]:
            filtered_mail = get_sender_mail(mails)
            return filtered_mail
        elif result == filters_list[2]:
            filtered_mail = get_mail_sender_and_subject(mails)
            return filtered_mail


def check_exit_program():
    user_input = input("Press enter to continue or insert 'off' to end the program.\n")
    
    if user_input == 'off':
        sys.exit()            
    else:
        return


def main():
    account = get_account()
    inbox = enter_application(account)
    mails = get_all_mail(inbox)

    while True:
        menu_selection = get_user_selection()

        # list all mail
        if menu_selection == '1':
            system('cls')
            result = select_type_mail(filters_list)           
            check_result(result, mails)
            check_exit_program()
            
        # create new folder
        elif menu_selection == '2':
            system('cls')            
            new_folder_name = get_new_folder_name()
            create_folder(inbox, new_folder_name)
            check_exit_program()

        # move mail to folder
        elif menu_selection == '3':
            system('cls')                              
            result = select_type_mail(filters_list)

            if check_result(result, mails) is not None:              
                folder_name = get_folder_name()
                move_to_folder(folder_name, filtered_mail, inbox)
    
            check_exit_program()

        # mark mail as read
        elif menu_selection == '4':
            system('cls')
            result = select_type_mail(filters_list)
            filtered_mail = check_result(result, mails)
            mark_mail_as_read(filtered_mail)    

            check_exit_program()            
        
        # delete mail
        elif menu_selection == '5':
            system('cls')
            pass

        # mark ALL as read
        elif menu_selection == '6':
            system('cls')
            pass

        # delete ALL mail
        elif menu_selection == '7':
            system('cls')
            pass

        # exit program withouth having to select an option on the menu
        elif menu_selection == '8':
            system('cls')
            check_exit_program()

        else:
            print("Wrong input. Try again. ")


## MAIN
if __name__ == '__main__':
    main()  