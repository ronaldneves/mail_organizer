import sys
import win32com.client as w32

type_list = ['subject', 'sender']

## FUNCTIONS
def 


def get_account():
    '''gets and returns the user account. must be str with a valid mail @outlook/@live/@hotmail.'''

    account = input("Insert your outlook/live/hotmail account: ")
    
    return account


def enter_application(account):
    '''takes the account from the user input to create the >inbox< object which contains all mail and returns it'''
    
    outlook = w32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    account = outlook.Folders[account]
    inbox = account.Folders['Caixa de Entrada']

    return inbox


def get_all_mail(inbox):
    '''retrieves all mail from the inbox and puts into a >mails< object'''

    mails = inbox.Items.Restrict("[ReceivedTime] >= '01/01/2007'")
    # items = list(items)[:200]

    return mails


def select_type_mail(type_list):
    '''takes the list of supported filters from the type_list and let the user chooses how he wants to filter his inbox.'''

    print("Insert the type of mail you want to filter from the list below: ")
    print(type_list)
    user_input = input("")

    if user_input in type_list:
        return user_input
    
    else:
        print(f"Type invalid or not supported. Supported filter types: {type_list}")    
        return None


def get_sender_mail(items):
    '''gets the items object to retrieve all the inbox's mail sent from the sender the user inputs and and prints them. then return the list of emails.''' 

    user_input = input("Insert the sender address you want to filter from your inbox: ")
    emails = [m for m in items if user_input in m.SenderEmailAddress.lower()]   
   
    for email in emails:
        print(email)

    return email


def get_mail_by_subject(items):
    '''gets the items object to retrieve all the inbox's mail with the user inputs on the subject and and prints them. then return the list of emails.''' 

    user_input = input("Insert the content on the mail's subject you want to filter from your inbox: ")
    print("")
    emails = [m for m in items if user_input in m.Subject.lower()]    

    for email in emails:
        print(email)

    return email


def main():
    while True:
        account = get_account()
        inbox = enter_application(account)
        mails = get_all_mail(inbox)
        result = select_type_mail(type_list)

        if result is not None:
            if result == type_list[0]:
                filtered_mail = get_mail_by_subject(mails)
            elif result == type_list[1]:
                filtered_mail = get_sender_mail(mails)
    
        user_input = input("Press enter to continue or insert 'off' to end the program.  ")

        if user_input == 'off':
            sys.exit()
        
        else:
            pass


## MAIN
if __name__ == '__main__':
    main()  