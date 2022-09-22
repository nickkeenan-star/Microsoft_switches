


from secrets import choice
from select import select


def outlookSwithes():
    howTo = ('To put outlook in safemod type: outlook.exe/safe ')
    switchList = ['0 - /a', '1 - /altvba otmfilename', '2 - /c messageclass', '3 - /checkclient', '4 - /cleanautocompletecache',
    '5 - /cleancategories', '6 - /cleanclientrules', '7 - /cleandmrecords', '8 - /cleanfinders ', '9 - /cleanfromaddresses',
    '10 - /cleanmailtipcache', '11 - /cleanreminders', '12 - /cleanroamedprefs', '13 - /cleanrules', '14 - /cleanserverrules',
    '15 - /cleansharing', '16 - /cleansniff', '17 - /cleansubscriptions', '18 - /cleanweather', '19 - /cleanviews', '20 - /embedding',
    '21 - /f msgfilename', '22 - /finder', '23 - /hol holfilename', '24 - /ical icsfilename', '25 - /importNK2', '26 - /importprf prffilename', '27 - /launchtraininghelp assetid', '28 - /m emailname',
    '29 - /noextensions', '30 - /nopreview', '31 - /p msgfilename', '32 - /profile profilename', '33 - /profiles', '34 - /promptimportprf', '35 - /recycle', '36 - /remigratecategories', '37 - /resetfolders', '38 - /resetfoldernames', '39 - /resetformregions',
    '40 - /resetnavpane', '41 - /resetquicksteps', '42 - /resetsearchcriteria', '43 - /resetsharedfolders', '44 - /resettodobar', '45 - /restore', '46 - /rpcdiag', '47 - /safe', '48 - /safe:1', '49 - /safe:3', '50 - /select foldername', '51 - /share \n       a - feed://URL/filename \n       b - /share stssync://URL \n       c - /share \n       d - web://URL/filename',
    '52 - /sniff', '52 - /t oftfilename', '53 - /v vcffilename', '54 - /vcal vcsfilename']
    print(howTo)
    for i in range(0, len(switchList)):

     print(switchList[i])
    
    

    infoSwitch = ['Creates an item with the specified file as an attachment.', 'Opens the VBA program specified in otmfilename, instead of %appdata%\microsoft\outlook\vbaproject.otm.', 'Creates a new item of the specified message class (Outlook forms or any other valid MAPI form).', 'Prompts for the default manager of e-mail, news, and contacts.', 'Removes all names and e-mail addresses from the Auto-Complete list. (Outlook 2013, 2016 only)', 'Deletes any custom category names that you have created. Restores categories to the default names.',
    'Starts Outlook and deletes client-based rules.', 'Deletes the Conversations Actions Table (CAT). CAT entries for a conversation thread usually expire 30 days after no activity. The command-line switch clears all conversation tagging, ignore, and moving rules immediately stopping any additional actions. (Outlook 2013, 2016 only)', 'Deletes the logging records saved when a manager or a delegate declines a meeting.', 'Resets all Search Folders in the Microsoft Exchange mailbox for only the first profile opened.', 'Removes all manually added From entries from the profile.',
    'Removes all MailTips from the cache. (Outlook 2013, 2016 only)', 'Clears and regenerates reminders.', 'All previous roamed preferences are deleted and copied again from the local settings on the computer where this switch is used. This includes the roaming settings for reminders, free/busy grid, working hours, calendar publishing, and RSS rules.', 'Starts Outlook and deletes client-based and server-based rules.Important If you have multiple or additional mailboxes in your Outlook profile, running the /cleanrules command line switch deletes the rules from all connected mailboxes. Therefore, it is recommended that you only run this command when your Outlook profile only contains the one, target mailbox.',
    'Starts Outlook and deletes server-based rules.', '	Removes all RSS, Internet Calendar, and SharePoint subscriptions from Account Settings, but leaves all the previously downloaded content on your computer.This is useful if you cannot delete one of these subscriptions within Outlook 2013.', 'Overrides the programmatic lockout that determines which of your computers (when you run Outlook at the same time) processes meeting items. The lockout process helps prevent duplicate reminder messages. This switch clears the lockout on the computer it is used. This enables Outlook to process meeting items.',
    'Deletes the subscription messages and properties for subscription features.', 'Removes city locations added to the Weather Bar.', 'Restores default views. All custom views you created are lost.', 'Used without command-line parameters for standard OLE co-create.', 'Opens the specified message file (.msg) or Microsoft Office saved search (.oss).', 'Opens the Advanced Find dialog box.', 'Opens the specified .hol file.', 'Opens the specified .ics file.', 'Imports the contents of an .nk2 file which contains the nickname list used by both the automatic name checking and Auto-Complete features.', 'Starts Outlook and opens/imports the defined MAPI profile (*.prf). If Outlook is already open, queues the profile to be imported on the next clean start.',
    'Opens a Help window with the Help topic specified in assetid displayed.', 'Provides a way for the user to add an e-mail name to the item. Only works together with the /c command-line parameter.', 'Both native and managed Component Object Model (COM) add-ins are turned off.', 'Starts Outlook with the Reading Pane off.', 'Prints the specified message (.msg).', 'Loads the specified profile. If your profile name contains a space, enclose the profile name in quotation marks (" ").', 'Opens the Choose Profile dialog box regardless of the Options setting on the Tools menu.', 'Same as /importprf except a prompt appears and the user can cancel the import.', 'Starts Outlook by using an existing Outlook window, if one exists. Used in combination with /explorer or /folder.', 'Starts Outlook and starts the following commands on the default mailbox: Upgrades colored For Follow Up flags to Outlook 2013 color categories. Upgrades calendar labels to Outlook 2013 color categories. Adds all categories used on non-mail items into the Master Category List',
    'Restores missing folders at the default delivery location', 'Resets default folder names (such as Inbox or Sent Items) to default names in the current Office user interface language.', 'Empties the form regions cache and reloads the form region definitions from the Windows registry.', 'Clears and regenerates the Folder Pane for the current profile.', 'Restores the default Quick Steps. All user-created Quick Steps are deleted.', 'Resets all Instant Search criteria so the default set of criteria is shown in each module.', 'Removes all shared folders from the Folder Pane.', 'Clears and regenerates the To-Do Bar task list for the current profile.', 'Attempts to open the same profile and folders that were open prior to an abnormal Outlook shutdown. (Outlook 2013, 2016 only)', 'Opens Outlook and displays the remote procedure call (RPC) connection status dialog box.', 'Starts Outlook without the Reading Pane or toolbar customizations. Both native and managed Component Object Model (COM) add-ins are turned off.', 'Starts Outlook with the Reading Pane off.', 'Both native and managed Component Object Model (COM) add-ins are turned off.', 
    'Starts Outlook and opens the specified folder in a new window. For example, to open Outlook and display the default calendar, use: "c:\program files\microsoft office\office15\outlook.exe" /select outlook:calendar.', 'https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6#Category=Outlook:~:text=Specifies%20a%20sharing%20URL%20to%20connect%20to%20Outlook.%20For%20example%2C%20use%20stssync%3A//URL%20to%20connect%20a%20SharePoint%20list%20to%20Outlook.', 'Starts Outlook, forces a detection of new meeting requests in the Inbox, and then adds them to the calendar.', 'Opens the specified .oft file.', 'Opens the specified .vcf file.', 'Opens the specified .vcs file.']
    
    

   
      
outlookSwithes()


