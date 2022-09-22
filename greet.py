


def greet():
    
    appFunction = ('This App is for swithches for Microsoft helps you fix shit! \nHave fun!')
    print(appFunction)
    win365servicelist = ['0 - Word', '1 - Excel', '2 - PowerPoint, Power Point Viewer', '3 - Outlook', '4 - Access']
    for i in range(0, len(win365servicelist)):
        print(win365servicelist[i])
    
    choice = input('Please choose what 365 service you are trying to fix:' )
    print(choice)

   
    
greet()