

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<center><h1>Search & Delete</h1></center>

<center><h2>eDiscovery using EWS & Graph APIs</h2></center>

</center>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# About Me

<img src="img/IRunicorn.jpg"
     alt="Markdown Monster icon"
     style="float: left; margin-right: 10px; width:400px;" />

<div>
    <div style="text-align: right"><h1>Josh Rickard</br>
                                        @MSAdministrator</h1>
    <h2><ul>
        <li>Blue Team</li>
        <li>DFIR</li>
        <li>Automate All the Things</li>
        <li>Open Sorcerer</li>
    </ul></h2>
</div>
</div>
    <br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<h2>I like long walks through binary trees, #Python & #PowerShell riding, holding hands with Microsoft, writing romantic experiences about code, git prune all the branches, and deep sea #phishing</h2>
<br>
<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<center><h1>Exchange Client Communications</h1></center>
    <div>
        <img src="img/exchange-1.jpg"
        alt="Markdown Monster icon"
        style="float: right; margin-top: 10px;" />
    </div>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


<center><h1>Exchange Web Services</h1></center>
<div>
    <div>
        <img src="img/exchange-3.png"
        alt="Markdown Monster icon"
        style="float: left; margin-right: 20px; height:900px;" />
    </div>
    <div>
        <h1>HTTP/S SOAP XML Service</br>
        </br>
        Basically, central access to all services/features of Microsoft Exchange
    </div>
</div>

<br>
<br>
<br><br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br><br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br><br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# py-ews

# Welcome to py-ews's documentation!


```

    .______   ____    ____       ___________    __    ____   _______.
    |   _  \  \   \  /   /      |   ____\   \  /  \  /   /  /       |
    |  |_)  |  \   \/   / ______|  |__   \   \/    \/   /  |   (----`
    |   ___/    \_    _/ |______|   __|   \            /    \   \    
    |  |          |  |          |  |____   \    /\    / .----)   |   
    | _|          |__|          |_______|   \__/  \__/  |_______/    
                                                                 


    A Python package to interact with Exchange Web Services
```


# **py-ews** is a cross platform python package to interact with both Exchange 2010 to 2019 on-premises and Exchange Online (Office 365).  This package will wrap all Exchange Web Service endpoints, but currently is focused on providing eDiscovery endpoints. 

<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Repository: https://github.com/swimlane/pyews
# Documentation: https://py-ews.readthedocs.io/en/latest/


<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br><br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br><br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Versions

## py-ews works with 2.7+ and 3.X versions of Python

# Installation

> Python's standard package manager is pip which uses packages from pypi

## OS X & Linux:

```
pip install py-ews
```

## Windows:

```
pip install py-ews
```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


# UserConfiguration object will be used for all calls to Exchange Web Services

## UserConfiguration object

```python
from pyews import UserConfiguration
```

## You have several options when it comes to authentication to EWS

# eDiscovery or Compliance Rights

## The most common is using an accounts username and password which has one of the following rights within your Exchange environment:

* ## **eDiscovery Rights**
* ## **Compliance Administrator**
* ## **eDiscovery Manager**
* ## **eDiscovery Administrator**
* ## **Organization Management**

```python
from pyews import UserConfiguraton

userconfig = UserConfiguration(
   'l337hacker@funtimes.onmicrosoft.com',
   'Password1234!'
)
```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Impersonation Rights

> ## Impersonation means you can act on behalf of a user/mailbox

* ## **ApplicationImpersonation** or **ms-exch-impersonation** must be given to the mailboxes/accounts that you have rights to impersonate

```python
from pyews import Impersonation
from pyews import UserConfiguraton

my_impersonation = Impersonation(principalname='first.last@company.com')
#Impersonation(primarysmtpaddress='first.last@company.com')
#Impersonation(smtpaddress='first.last@company.com')


userconfig = UserConfiguration(
   'l337hacker@funtimes.onmicrosoft.com',
   'Password1234!',
   impersonation=my_impersonation
)
```
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Impersonation DeepDive

```python
class Impersonation(object):
    '''The Impersonation class is used when you want to impersonate a user.  You must access rights to impersonate a specific user within your Exchange environment.
    
    Example:
        Below are examples of the data inputs expected for all parameters.
        
        .. code-block:: python

           Impersonation(principalname='first.last@company.com')
           Impersonation(primarysmtpaddress='first.last@company.com')
           Impersonation(smtpaddress='first.last@company.com')

    Args:
        principalname (str, optional): The PrincipalName of the account you want to impersonate
        sid (str, optional): The SID of the account you want to impersonate
        primarysmtpaddress (str, optional): The PrimarySmtpAddress of the account you want to impersonate
        smtpaddress (bool, optional): The SmtpAddress of the account you want to impersonate
    
    Raises:
        AttributeError: This will raise when you call this class but do not provide at least 1 parameter
    '''

    def __init__(self, principalname=None, sid=None, primarysmtpaddress=None, smtpaddress=None):

        if principalname:
            self.impersonation_type = 'PrincipalName'
            self.impersonation_value = principalname
        elif sid:
            self.impersonation_type = 'SID'
            self.impersonation_value = sid
        elif primarysmtpaddress:
            self.impersonation_type = 'PrimarySmtpAddress'
            self.impersonation_value = primarysmtpaddress
        elif smtpaddress:
            self.impersonation_type = 'SmtpAddress'
            self.impersonation_value = smtpaddress
        else:
            raise AttributeError('By setting impersonation to true you must provide either a PrincipalName, SID, PrimarySmtpAddress, or SmtpAddress')

        self.header = self._create_impersonation_header()

    def _create_impersonation_header(self):
        return '''<soap:Header>
  <t:ExchangeImpersonation>
    <t:ConnectingSID>
      <t:{start_type}>{value}</t:{end_type}>
    </t:ConnectingSID>
  </t:ExchangeImpersonation>
</soap:Header>'''.format(start_type=self.impersonation_type, value=self.impersonation_value, end_type=self.impersonation_type)

```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


# If you would like to use impersonation, you first need to create an Impersonation object and pass that into the UserConfiguration class when creating a user configuration object.

```python
from pyews import UserConfiguration

impersonation = Impersonation(primarysmtpaddress='first.last@company.com')

userConfig = UserConfiguration(
    'first.last@company.com',
    'mypassword123',
    impersonation=impersonation
)
```


<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


# Autodiscover

## Autodiscover is a service endpoint that is used to determine the Exchange Web Services URL/Endpoint

# Microsoft recommends taking the following approach to discover Autodiscover endpoints

* ## Lookup Service Connection Point (SCP) Objects in Active Directory
* ## Generate potential locations based on the users email address

<br>
<br>
<br>
<br>

# py-ews does not currently utilize SCP objects in Active Directory

> ## py-ews does not do this as it would increase the size of this package significantly

<br>
<br>
<br>
<br>

# py-ews generates a list of potential endpoints based on the provided users email address

## Example: first.last@letsautomate.it

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>

# By default py-ews will use each of the following URLs to automatically connect you via Autodiscover

```
https://letsautomate.it/autodiscover/autodiscover.svc
https://autodiscover.letsautomate.it/autodiscover/autodiscover.svc
http://letsautomate.it/autodiscover/autodiscover.svc
http://autodiscover.letsautomate.it/autodiscover/autodiscover.svc

https://outlook.office365.com/autodiscover/autodiscover.svc
```



## Example of using Autodiscover

```python
from pyews import UserConfiguration

userconfig = UserConfiguration(
   'l337hacker@funtimes.onmicrosoft.com',
   'Password1234!'
)

# you can print properties on the useconfig object if needed

#print(userconfig.autodiscover)
#print(userconfig.configuration)
#print(userconfig.credentials.password)
#print(userconfig.credentials.email_address)
#print(userconfig.ewsUrl)
#print(userconfig.exchangeVersion)
#print(userconfig.raw_soap)
```
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Autodiscover DeepDive

```xml
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover"      
               xmlns:wsa="http://www.w3.org/2005/08/addressing" 
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"      
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
  <soap:Header>
    <a:RequestedServerVersion>{version}</a:RequestedServerVersion>
    <wsa:Action>http://schemas.microsoft.com/exchange/2010/Autodiscover/Autodiscover/GetUserSettings</wsa:Action>
    <wsa:To>{to}</wsa:To>
    <t:ExchangeImpersonation>
        <t:ConnectingSID>
            <t:{start_type}>{value}</t:{end_type}>
        </t:ConnectingSID>
    </t:ExchangeImpersonation>
  </soap:Header>
  <soap:Body>
    <a:GetUserSettingsRequestMessage xmlns:a="http://schemas.microsoft.com/exchange/2010/Autodiscover">
      <a:Request>
        <a:Users>
          <a:User>
            <a:Mailbox>{mailbox}</a:Mailbox>
          </a:User>
        </a:Users>
        <a:RequestedSettings>
          <a:Setting>InternalEwsUrl</a:Setting>
          <a:Setting>ExternalEwsUrl</a:Setting>
          <a:Setting>UserDisplayName</a:Setting>
          <a:Setting>UserDN</a:Setting>
          <a:Setting>UserDeploymentId</a:Setting>
          <a:Setting>InternalMailboxServer</a:Setting>
          <a:Setting>MailboxDN</a:Setting>
          <a:Setting>ActiveDirectoryServer</a:Setting>
          <a:Setting>CasVersion</a:Setting>
          <a:Setting>EwsSupportedSchemas</a:Setting>
        </a:RequestedSettings>
      </a:Request>
    </a:GetUserSettingsRequestMessage>
  </soap:Body>
</soap:Envelope>
```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Additional Options available when creating a UserConfiguration object

## Provide specific version of Exchange

```python
from pyews import UserConfiguration

userConfig = UserConfiguration(
    'first.last@company.com',
    'mypassword123',
    exchangeVersion='Office365'
)

# options: ['Office365', 'Exchange2019', 'Exchange2016', 'Exchange2013_SP1', 'Exchange2013', 'Exchange2010_SP2', 'Exchange2010_SP1', 'Exchange2010']
```
<br>
<br>
<br>
<br>

# If you do not wish to use Autodiscover then you can tell UserConfiguration to not use it by setting autodiscover to False and provide the ewsUrl instead

```python
from pyews import UserConfiguration

userConfig = UserConfiguration(
    'first.last@company.com',
    'mypassword123',
    autodiscover=False,
    ewsUrl='https://outlook.office365.com/EWS/Exchange.asmx'
)
```

<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Currently Available Service Endpoints

* # DeleteItem
    * ## Delete a mail item (message) from a users mailbox
* # GetInboxRules
    * ## Retrieve mailbox/inbox rules from a users mailbox
* # GetSearchableMailboxes
    * ## Retrieve all mailboxes that you have access to search
* # ResolveNames
    * ## Take a users identify and resolve it to retrieve details about that user
* # SearchMailboxes
    * ## Search a mailbox 
<br>
<br>
<br>

# You can also extend py-ews fairly easily

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Typical workflow

* # Retrieve all mailboxes

```python
from pyews import GetSearchableMailboxes

# get searchable mailboxes based on your accounts permissions
referenceid_list = []
for mailbox in GetSearchableMailboxes(userconfig).response:
    referenceid_list.append(mailbox['ReferenceId'])
```
> ## NOTE: GetSearchableMailboxes will automatically expand any found groups and their contained mailboxes
<br>

* # Now let's search all mailboxes

```python
from pyews import SearchMailboxes

# let's search all the referenceid_list items
messages_found = []
for search in SearchMailboxes('subject:account', userconfig, referenceid_list).response:
    messages_found.append(search['MessageId'])
    # we can print the results first if we want
    print(search['Subject'])
    print(search['MessageId'])
    print(search['Sender'])
    print(search['ToRecipients'])
    print(search['CreatedTime'])
    print(search['ReceivedTime'])
    #etc.
```

* # Deleting a specific message by MessageId


```python
from pyews import DeleteItem

# if we wanted to now delete a specific message then we would call the DeleteItem
# class like this but we can also pass in the entire messages_found list
deleted_message_response = DeleteItem(messages_found[2], userconfig).response

print(deleted_message_response)

```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


# Running GetSearchableMailboxes and printing out properties

<img src="img/get_searchable_mailboxes.gif"
        alt="Markdown Monster icon"
        style="float: center; margin-right: 10px;" />


<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# SearchMailboxes

> ## Given a single or a list of mailbox ReferenceId's, we can perform a search

```python
from pyews import SearchMailboxes

# let's search all the referenceid_list items
messages_found = []
for search in SearchMailboxes('subject:account', userconfig, referenceid_list).response:
    # We need the MessageId from any email we may want to delete
    messages_found.append(search['MessageId'])
    # we can print the results first if we want
    print(search['Subject'])
    print(search['MessageId'])
    print(search['Sender'])
    print(search['ToRecipients'])
    print(search['CreatedTime'])
    print(search['ReceivedTime'])
    #etc.
```

# We provide a search query, as well as our UserConfiguration Object and a single or list of ReferenceId's to SearchMailboxes


<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Running SearchMailboxes and printing all properties

<img src="img/mail_item_properties.gif"
        alt="Markdown Monster icon"
        style="float: center; margin-right: 10px;" />


<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Exchange Web Services Advanced Query Syntax (AQS)

## Besides just searching for `subject:account` in our example on the previous slide you have the following `AQS` options


<br>
<br>
<br>
<div style="width: 100%; overflow: hidden;">
    <div style="width: 800px; float: left;"> <img src="img/search_query_options.png" style="float: center; margin-right: 10px;" /> </div>
    <div style="margin-left: 820px;"> 
        <h1><b>subject:account AND body:click AND received:today</br></br></br>
               subject:account AND body:click AND attachment:fax</br></br></br>
               from:"Mr. Robot" OR from:"Josh" AND subject:(raise upgrade warning expire) </br></br></br>
               subject:"Office 365" NOT from:"IT Department"</br></br></br>
               body:http* AND received:today</br></br></br>
               attachment:* AND received:today
    </h1> </div>
</div>


<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# DeleteItem

> ## We have a few options when we want to delete an email

* ## HardDelete (An item is permanently removed from the store)
* ## SoftDelete (An item is moved to the dumpster if the dumpster is enabled)
* ## MoveToDeletedItems (An item is moved to the Deleted Items folder) (Default)


```python
from pyews import DeleteItem

# if we wanted to now delete a specific message then we would call the DeleteItem
# class like this but we can also pass in the entire messages_found list
deleted_message_response = DeleteItem(messages_found[2], userconfig).response

# messages_found[2] translates to a single MessageId (AAMkAGZjOTlkOWExLTM2MDEtNGI3MS04ZDJiLTllNzgwNDQxMThmMABGAAAAAABdQG8UG7qjTKf0wCVbqLyMBwC6DuFzUH4qRojG/OZVoLCfAAAAAAEMAAC6DuFzUH4qRojG/OZVoLCfAAAu4Y9UAAA=)


print(deleted_message_response)

```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

```
  _______ .______          ___      .______    __    __   __       _______. __    __  
 /  _____||   _  \        /   \     |   _  \  |  |  |  | |  |     /       ||  |  |  | 
|  |  __  |  |_)  |      /  ^  \    |  |_)  | |  |__|  | |  |    |   (----`|  |__|  | 
|  | |_ | |      /      /  /_\  \   |   ___/  |   __   | |  |     \   \    |   __   | 
|  |__| | |  |\  \----./  _____  \  |  |      |  |  |  | |  | .----)   |   |  |  |  | 
 \______| | _| `._____/__/     \__\ | _|      |__|  |__| |__| |_______/    |__|  |__| 
                                                                                      

A Python package to search & delete messages from mailboxes in Office 365 using Microsoft Graph API
```

## Current Features

    * Searching
        * Create new Search
        * Update a Search
        * Get Search Folder
        * Get Search messages
        * Delete Search
    * Deleting
        * Delete a Message
    * Mailbox Rules
        * List Mailbox Rules
    * Users
        * Return a list of email addresses in your Azure AD Tenant
    * MailFolder
        * Create MailFolder
        * Move message to a MailFolder
        * List messages in a MailFolder

## Installation

OS X & Linux:

```bash
pip install graphish
```

Windows:

```bash
pip install graphish
```
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# OAuth2 Authentication Types & Grant Flows

|Application Scenarios|Application Type|Implicit Grant Flow|Auth Code Grant Flow|On-Behalf-Of Flow|Client Credentials Grant|
|:---------------------|:---------------------|:---------------------:|:---------------------:|:---------------------:|:---------------------:|
|Single-Page Application|Single-Page Apps|X||||
|Web Browser to Web Application|Web Apps|X|X|||
|Native Application to Web API|Mobile & Native Apps|X|X|X||
|Web Application to Web API|Web APIs|||X|X|
|Daemon or Service to Web API|Daemons & Server-Side Apps||||x|

<br>
<br>

# Summary

## Delegated Authentication (SPA, Web Apps, Mobile/Native Apps, etc.) == User Interaction Authentication Required

<img src="img/Part-3_Permissions-requested.png"
        alt="Markdown Monster icon"
        style="float: center; margin-right: 10px; width:300px;" />



## Application Authentication (Daemon/Service) == No User Interaction Authentication Required 

<br>
<br>

> If you are wanting more info about OAuth2 Authentication check out this blog post series I wrote

* https://swimlane.com/blog/microsoft-oauth2-implementation-1/
* https://swimlane.com/blog/microsoft-oauth2-implementation-2/
* https://swimlane.com/blog/microsoft-oauth2-implementation-3/

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Delegated Authentication Example

```python
from graphish import GraphConnector

connector = GraphConnector(
    clientId='14b8e5asd-c5a2-4ee7-af26-53461f121eed',       # you applications clientId
    clientSecret='OdhG1hXb*UB/ho]A?0ZCci13KMflsHDy',        # your applications clientSecret
    tenantId='c1141d00-072f-1eb9-2526-12802571dd41',        # your applications Azure Tenant ID
    userPrincipalName='first.last@myorg.onmicrosoft.com',   # the user's mailbox you want to search
    password='somepassword'                                 # password of your normal or admin account
)
```

# Application Authentication Example

```python
from graphish import GraphConnector

# For backend / client_credential auth flow just supply the following
connector = GraphConnector(
    clientId='14b8e5asd-c5a2-4ee7-af26-53461f121eed',       # you applications clientId
    clientSecret='OdhG1hXb*UB/ho]A?0ZCci13KMflsHDy',        # your applications clientSecret
    tenantId='c1141d00-072f-1eb9-2526-12802571dd41',        # your applications Azure Tenant ID
)
```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<center><h1>Why Microsoft Graph API is different</h1></center>

<br>
<br>

# Microsoft Graph API creates a "Search" folder in a users mailbox that is hidden to them and can only be accessed by users with appropriate permissions

<br>
<br>

# Once you then configure this "Search" folder to populate based on your search criteria (or AQL I mentioned previously).  

<br>
<br>

# Any email matching this criteria will copied to this folder and then you can retrieve it

<br>
<br>

# This also means, we can update, modify, or delete our search rules for a specific "Search" folder

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


# Creating a new Search


> **NOTE: If you DO NOT pass in a userPrincipalName then Search will attempt to search all mailboxes in your Azure AD tenant!**

```python
from graphish import Search

search = Search(connector)

new_search = search.create(
    searchFolderName='Phishing Search',
    sourceFolder='inbox',
    filterQuery="contains(subject, 'EXPIRES')"
)


# get all the messages in your search folder

messages = search.messages()
for message in messages:
    print(message) # Returns all attributes from a message
    print(message['id']) # returns the message ID
    print(message['headers']) # Returns the RFC822 headers of the message


# Deleting a message

#

from graphish import Delete

delete = Delete(connector, userPrincipalName='some.user@company.com')

# delete a specific message ID from users mailbox

delete.delete(message['id'])

# delete a specific message from a search (or any) folder in the users mailbox

delete.delete_search_message('Phishing Search', message['id'])

```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


```python
class Delete(object):

    def __init__(self, graphConnector, userPrincipalName='me'):
        """The Delete class deletes a message from a users mailbox.  

            NOTE: Please note that this will do a soft delete and it is only recoverable via "Recoverable Items" feature on a mailbox.
                  This also means that you will not see it in the users delete items folder.
        
        Args:
            graphConnector (GraphConnector): A generated GraphConnector object
            verify_ssl (bool, optional): Whether to verify SSL or not. Defaults to True.
            userPrincipalName (str, optional): Defaults to the current user, but can be any user defined or provided in this parameter. Defaults to 'me'.
        """
        self.connector = graphConnector
        if userPrincipalName is not 'me':
            self.user = 'users/{}'.format(userPrincipalName)
        else:
            self.user = userPrincipalName

    def delete(self, messageId):
        uri = '{}/messages/{}'.format(self.user, messageId)
        response = self.connector.invoke('DELETE', uri)
        if response.status_code is '204':
            return True
        return False

    def delete_search_message(self, mailFolder, messageId):
        uri = '{}/mailFolders/{}/messages/{}'.format(self.user, mailFolder, messageId)
        response = self.connector.invoke('DELETE', uri)
        if response.status_code is '204':
            return True
        return False
```

<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

# Updating a search

If you wanted to make changes to a search performed you can update the search folder and individual criteria like the name of the search folder, the sourceFolder (root to search), or the filterQuery itself:

```python
# update your search folder property's
search.update(
    searchFolderName='Some Phishing Search',
    sourceFolder='inbox',
    filterQuery="contains(subject, 'EXPIRES!')"
)
```

# Deleting a search

You can also delete a search performed by using the `delete` method:

```python
# delete the current search folder
search.delete()
```



<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>


<center><h1>FIN</h1></center>

<br>
<br>

# py-ews: [https://py-ews.readthedocs.io/en/latest/](https://py-ews.readthedocs.io/en/latest/)
# graphish: [https://github.com/swimlane/graphish](https://github.com/swimlane/graphish)

<br>
<br>

# Twitter: [@MSAdministrator](https://twitter.com/msadministrator)
# GitHub: [https://github.com/MSAdministrator](https://github.com/MSAdministrator)
# Blog: [https://letsautomate.it](https://letsautomate.it)

