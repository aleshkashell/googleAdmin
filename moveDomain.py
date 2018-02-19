from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from apiclient import discovery
import httplib2
import os
import sys
import time #sleep
import json
from datetime import datetime
import re

# Email of the Service Account
SERVICE_ACCOUNT_EMAIL = 'your_account_name.iam.gserviceaccount.com'

# Path to the Service Account8b1ec96's Private Key file
SERVICE_ACCOUNT_PKCS12_FILE_PATH = os.path.expanduser('~/Documents/.secret/service.p12')
adminEmail = 'adminEmail@yourdomain.ru'
mainDomain = 'yourdomain.com'
siteName = 'yourdomain'
newDomain = 'com'
currentDomain = 'ru'

def get_credentials(email):
    #Авторизация и делегация на пользователя e-mail
    credentials = ServiceAccountCredentials.from_p12_keyfile(
        SERVICE_ACCOUNT_EMAIL,
        SERVICE_ACCOUNT_PKCS12_FILE_PATH,
        'notasecret',
        scopes=['https://www.googleapis.com/auth/admin.directory.user', 'https://www.googleapis.com/auth/gmail.settings.sharing',
                'https://www.googleapis.com/auth/gmail.settings.basic', 'https://mail.google.com/', 'https://www.googleapis.com/auth/activity',
                'https://www.googleapis.com/auth/drive.metadata.readonly', 'https://www.googleapis.com/auth/admin.directory.group'])
    delegate_credentials = credentials.create_delegated(email)
    return delegate_credentials.authorize(httplib2.Http())
def backUpInfo(filename=''):
    path = os.path.expanduser('~/Documents/backup/')
    prefix = path+'backup'
    if not (filename == ''):
        prefix = path + filename
    #backupUsers
    with open(prefix + 'Users.json', 'w') as file:
        userInfo = getAllInfoUsers()
        file.write(json.dumps(userInfo, indent=4))
    print(prefix + 'Users.json done')
    #backupGroups
    with open(prefix + 'Groups.json', 'w') as file:
        groupInfo = getAllInfoGroups()
        file.write(json.dumps(groupInfo, indent=4))
    print(prefix + 'Groups.json done')
    #backupUsersInGroups
    with open(prefix + 'UsersInGroups.json', 'w') as file:
        groupInfo = getAllUsersInGroups()
        file.write(json.dumps(groupInfo, indent=4))
    print(prefix + 'UsersInGroups.json done')
    #backupUsersAliases
    with open(prefix + 'UsersAliases.json', 'w') as file:
        info = getUsersAliases()
        file.write(json.dumps(info, indent=4))
    print(prefix + 'UsersAliases.json done')
    #backupGroupsAliases
    with open(prefix + 'GroupsAliases.json', 'w') as file:
        info = getGroupAliases()
        file.write(json.dumps(info, indent=4))
    print(prefix + 'GroupsAliases.json done')
def changeMailDomain(email, domain):
    if not '@' + siteName + '.' + domain in email:
        return email.split('@')[0] + '@' + siteName + '.' + domain
    return email
def createAliasDomain(email, domain, isDefault=True):
    newEmail = changeMailDomain(email, domain)
    sendAsResource = {"sendAsEmail": newEmail,
                      "isDefault": isDefault,
                      "replyToAddress": newEmail,
                      "displyaName": getFullname(email),
                      "isPrimary": False,
                      "treatAsAlias": False
                      }
    http = get_credentials(email)
    GMAIL = discovery.build('gmail', 'v1', http=http)
    return GMAIL.users().settings().sendAs().create(userId='me', body=sendAsResource).execute()
def createAliasNewName(email, alias, isDefault=False):
    body = {'alias': alias}
    http = get_credentials(adminEmail)
    service = discovery.build('admin', 'directory_v1', http=http)
    try:
        responce = service.users().aliases().insert(userKey=email, body=body).execute()
        return responce
    except Exception as e:
        return (sys.exc_info()[0], ":", e)
def createAliasNewNameGroup(email, alias):
    body = {'alias': alias}
    http = get_credentials(adminEmail)
    service = discovery.build('admin', 'directory_v1', http=http)
    try:
        responce = service.groups().aliases().insert(groupKey=email, body=body).execute()
        return responce
    except Exception as e:
        return (sys.exc_info()[0], ':', e)
def createAliasCom(email):
    emailCom = changeMailDomain(email=email, domain='com')
    sendAsResource = {"sendAsEmail": emailCom,
                      "isDefault": True,
                      "replyToAddress": emailCom,
                      "displayName": getFullname(email),
                      "isPrimary": False,
                      "treatAsAlias": False
                      }
    print(sendAsResource)
    http = get_credentials(email)
    GMAIL = discovery.build('gmail', 'v1', http=http)
    response = GMAIL.users().settings().sendAs().create(userId='me', body=sendAsResource).execute()
    print(response)
def createSignature(post, name, surname='', telephone='', skype=''):
    template = ('<div style="font-family:helvetica,arial;font-size:13px">Regards,</div>'
            '<div style="font-family:helvetica,arial;font-size:13px"><b>')
    if surname == '':
        template = template + name
    else:
        template = template + name +' ' + surname
    template = template + '</b><div style="font-family:helvetica,arial;font-size:13px"><i>' + post + '</i>'
    if not skype == '':
        template = template + '<div style="font-family:helvetica,arial;font-size:13px">Skype: ' + skype + '</div>'
    if not telephone == '':
        template = template +'<div>Tel: <a dir="ltr" href="tel:+%207%20916%207183639" target="_blank">' + telephone + '</a></div>'
    template += '<div><img src="https://ci4.googleusercontent.com/proxy/vsSFXEeaHOwP-g7pKbqB00dUtC1wwA1I_iG1KWsVXUfKDQg6Vvasf9u2LXsT1QYZI1YruapYZg=s0-d-e1-ft#http://cdn.pixapi.net/pixonic.png" width="250" class="CToWUd"></div><div><span style="color:rgb(0,0,0);font-size:9px">This e-mail and the attachments, if any, are confidential and may be legally privileged, and are for the sole use of the intended recipient. If you are not the intended recipient of this e-mail or any part of it, please delete it from your computer.</span></div></div></div>'
    return {'signature': template}
def createUserFull(email, firstname, secondname, post, department, telephone, skype):
    print("Create user " + email)
    create_user(email=email, firstname=firstname, secondname=secondname)
    print("Wait 15s")
    time.sleep(15)
    print("Update information about user")
    updateInformation(userEmail=email, post=post, department=department, telephone=telephone)
    print("Create signature")
    signature = createSignature(post=post, name=firstname, surname=secondname, telephone=telephone, skype=skype)
    print('Edit signature for ru')
    editSignature(email=email, signature=signature)
    print("wait 1 minute")
    time.sleep(60)
    print("Create alias com")
    createAliasCom(email)
    time.sleep(15)
    print('Edit signature for com ' + changeMailDomain(email=email, domain='com')(email))
    editSignature(email=changeMailDomain(email=email, domain='com'), signature=signature)
def create_user(email, firstname, secondname):
    http = get_credentials(adminEmail)
    service = discovery.build('admin', 'directory_v1', http=http)
    userinfo = {'primaryEmail': email,
        'name': { 'givenName': firstname, 'familyName': secondname },
        'password': 'Zz123456',}
    service.users().insert(body=userinfo).execute()
def editSignature(email, signature):
    http = get_credentials(email)
    GMAIL = discovery.build('gmail', 'v1', http=http)
    print(signature)
    response = GMAIL.users().settings().sendAs().patch(userId='me', sendAsEmail=email, body=signature).execute()
    print(response)
def getAliasOfUser(usermail):
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().aliases().list(userKey=usermail).execute()
    result = []
    try:
        for a in data['aliases']:
            result.append(a['alias'])
    except(KeyError):
        pass
    return result
def getGroupAliases():
    groups = getAllInfoGroups()
    root = {}
    for group in groups:
        root[group] = {}
        try:
            root[group]['aliases'] = groups[group]['aliases']
        except (KeyError):
            pass
        try:
            root[group]['nonEditableAliases'] = groups[group]['nonEditableAliases']
        except (KeyError):
            pass
    return root
def getSignature(email):
    http = get_credentials(email)
    GMAIL = discovery.build('gmail', 'v1', http=http)
    response = GMAIL.users().settings().sendAs().get(sendAsEmail=email, userId='me').execute()
    return response['signature']
def getFullname(email):
    return getUser(email)['name']['fullName']
def getPost(email):
    return getUser(email)['organizations'][0]['title']
def getUsers():
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().list(domain=mainDomain).execute()
    result = []
    for a in data['users']:
        if 'organizations' in a:
            result.append([a['primaryEmail'], a['name']['fullName'], a['organizations'][0]['title'], a['organizations'][0]['department']])
        else:
            result.append([a['primaryEmail'], a['name']['fullName'], "", ""])
    while 'nextPageToken' in data:
        data = service.users().list(domain=mainDomain, pageToken=data['nextPageToken']).execute()
        for a in data['users']:
            if 'organizations' in a:
                result.append([a['primaryEmail'], a['name']['fullName'], a['organizations'][0]['title'], a['organizations'][0]['department']])
            else:
                result.append([a['primaryEmail'], a['name']['fullName'], "", ""])
    return result
def getUsersAliases():
    users = getAllInfoUsers()
    root = {}
    for user in users:
        root[user] = {}
        try:
            root[user]['aliases'] = users[user]['aliases']
        except(KeyError):
            pass
        try:
            root[user]['nonEditableAliases'] = users[user]['nonEditableAliases']
        except(KeyError):
            pass
    return root
def getAllInfoUsers():
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().list(domain=mainDomain).execute()
    result = {}
    for a in data['users']:
        result[a['primaryEmail']] = a
        #result.append(a)
    while 'nextPageToken' in data:
        data = service.users().list(domain=mainDomain, pageToken=data['nextPageToken']).execute()
        for a in data['users']:
            result[a['primaryEmail']] = a
    return result
def getAllInfoGroups():
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.groups().list(domain=mainDomain).execute()
    result = {}
    for a in data['groups']:
        result[a['email']] = a
    while 'nextPageToken' in data:
        data = service.groups().list(domain=mainDomain, pageToken=data['nextPageToken']).execute()
        for a in data['groups']:
            result[a['email']] = a
    return result
def getMembersOfGroup(groupName):
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.members().list(groupKey=groupName).execute()
    result = {}
    try:
        for a in data['members']:
            try:
                result[a['email']] = a
            except (KeyError):
                pass
    except(KeyError):
        pass
    while 'nextPageToken' in data:
        data = service.members().list(groupKey=groupName, pageToken=data['nextPageToken']).execute()
        for a in data['members']:
            result[a['email']] = a
    return result
def getAllUsersInGroups():
    result = {}
    for groupName in getAllInfoGroups():
        print(groupName)
        #result[groupName['email']] = groupName['email']
        result[groupName] = getMembersOfGroup(groupName)
    return result
def getUser(email):
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().get(userKey=email).execute()
    return data
def loadJsonFile(filename):
    with open(filename, 'r') as file:
        result = json.load(file)
    return result
def mailToCom(email):
    if not ('@' + siteName + '.com') in email:
        return email.split('@')[0] + '@' + siteName + '.com'
    return email
def moveUsers(domain):
    # Получаем список всех пользователей
    users = getUsers()
    # Переименовываем всех пользователей, кроме adminEmail
    #print(json.dumps(users, indent=4))
    for a in users:
        if adminEmail in a[0]: continue
        try:
            renameUser(a[0], domain)
            print('Done: ' + a[0])
        except Exception as e:
            print(sys.exc_info()[0], ":", e)
            print('Failed: ' + a[0])
    print("Users done!!!")
def moveGroups(domain): #domain='com'
    #Получаем список групп
    groups = getAllInfoGroups()
    for group in groups:
        try:
           renameGroup(groups[group]['email'], domain)
           print('Done: ' + groups[group]['email'])
        except Exception as e:
           print(sys.exc_info()[0], ":", e)
           print('Failed: ' + groups[group]['email'])
    print("Groups done!!!")
def moveAll(domain):
    moveUsers(domain)
    moveGroups(domain)
def renameGroup(group, domain):
    patch = {'email': changeMailDomain(group, domain)}
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    service.groups().update(groupKey=group, body=patch).execute()
def renameUser(email, domain):
    patch = {'primaryEmail': changeMailDomain(email=email, domain=domain)}
    #print(json.dumps(patch, indent=4))
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    #print(email, patch)
    service.users().update(userKey=email, body=patch).execute()
def restoreAliasGroups(filename):
    groups = loadJsonFile(filename)
    result = {}
    for group in groups:
        try:
            aliases = groups[group]['aliases']
            if (len(aliases) > 0):
                result[group] = aliases
        except:
            pass
    aliases = result
    flag = True
    for group in aliases:
        for alias in aliases[group]:
            try:
                createAliasNewNameGroup(email=group, alias=changeMailDomain(email=alias, domain=newDomain))
                print("Done:", alias)
            except Exception as e:
                print("Failed:", alias)
                print(sys.exc_info()[0], ":", e)
                flag = False
    if(flag):
        print("All groups' aliases have been restored")
    else:
        print("We have a problem with groups")
def restoreAliasUsers(filename):
    users = loadJsonFile(filename)
    result = {}
    for user in users:
        try:
            aliases = users[user]['aliases']
            if (len(aliases) > 0):
                result[user] = aliases
        except:
            pass
    aliases = result
    flag = True
    for user in aliases:
        for alias in aliases[user]:
            try:
                createAliasNewName(email=user, alias=changeMailDomain(alias, newDomain))
                print("Done: ", alias)
            except Exception as e:
                print("Failed: ", alias)
                print(sys.exc_info()[0], ":", e)
                flag = False
    if(flag):
        print("All users' aliases have been restored")
    else:
        print("We have a problem with users")
def saveOutputToFile(lines, filename):
    with open(filename, 'w') as file:
        for a in lines:
            file.write(str(a))
            file.write('\n')
def updateInformation(userEmail, post, department, telephone):

    patch = {'organizations': [{'title': post, 'primary': True, 'customType': '',
        'department': department, 'description': '', 'costCenter': ''}],
        'phones': [{'value': telephone, 'type': 'home'}]}

    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    service.users().update(userKey=userEmail, body=patch).execute()

def updateAllUserPost():
    users = getPixTeam.getPixTeam()
    for email, post, department in users:
        result = updateInformation(adminEmail=adminEmail, userEmail=email, post=post, department=department)
        print(email, post, department)
        print(result)
    print("Done")
def updateSignatureFromGoogle(email): #For 'ru' & 'com'
    fullName = getUser(email=email)['name']['fullName']
    post = getUser(email=email)['organizations'][0]['title']
    signature = createSignature(name=fullName, post=post)
    #Update ru
    try:
        editSignature(email=changeMailDomain(email=email, domain='ru'), signature=signature)
    except Exception as e:
        print(changeMailDomain(email=email, domain='ru'), '\t', sys.exc_info()[0], ":", e)
    #Update com
    try:
        editSignature(email=changeMailDomain(email=email, domain='com'), signature=signature)
    except Exception as e:
        print(changeMailDomain(email=email, domain='com'), '\t', sys.exc_info()[0], ":", e)


def main():
#Снять следующий комментарий перед восстановлением из бекапа
#"""
    now = datetime.now()
    hFile = str(now.month) + str(now.day)
    global adminEmail # Needed to modify global copy of globvar
    #Делаем backup
    print("Делаем backup")
    backUpInfo(siteName + hFile)
    #Делаем миграцию в новый домен
    print("Делаем миграцию в новый номен")
    moveAll(newDomain)
    #Мигририум собственный аккаунт
    print("Мигрируем собственный аккаунт")
    renameUser(adminEmail, newDomain)
    #
    print("Переименовываем adminEmail")
    adminEmail = changeMailDomain(adminEmail, newDomain)
    print(adminEmail)

    #Ожидаем завершения операции смены главного домена
    #Авторизуем api в новом домене
    """
    #restore alias
    path = os.path.expanduser('~/Documents/backup/')
    usersAliases = path + siteName + hFile + "UsersAliases.json"
    print("Восстанавливаем пользовательские aliases")
    restoreAliasUsers(usersAliases)
    groupsAliases = path + siteName + hFile + "GroupsAliases.json"
    print("Восстанавливаем групповые aliases")
    restoreAliasGroups(groupsAliases)
#"""

if __name__ == '__main__':
    main()