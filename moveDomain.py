#!/usr/bin/env python
# -*- mode:python; coding:utf-8; -*-


import os
import sys
import time
import json
from datetime import datetime
import re

import httplib2
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from apiclient import discovery

####### Поправьте данные в соответствии с вашим случаем
####### Please edit config data for your case
# Email of the Service Account
SERVICE_ACCOUNT_EMAIL = 'your_account_name.iam.gserviceaccount.com'

# Path to the Service Account8b1ec96's Private Key file
SERVICE_ACCOUNT_PKCS12_FILE_PATH = os.path.expanduser('~/Documents/.secret/service.p12')
adminEmail = 'adminEmail@yourdomain.ru'
mainDomain = 'yourdomain.ru'
siteName = 'yourdomain'
newDomain = 'com'
currentDomain = 'ru'
##########################################################


def get_credentials(email):
    # Авторизация и делегация на пользователя e-mail
    credentials = ServiceAccountCredentials.from_p12_keyfile(
        SERVICE_ACCOUNT_EMAIL,
        SERVICE_ACCOUNT_PKCS12_FILE_PATH,
        'notasecret',
        scopes=['https://www.googleapis.com/auth/admin.directory.user',
                'https://www.googleapis.com/auth/gmail.settings.sharing',
                'https://www.googleapis.com/auth/gmail.settings.basic',
                'https://mail.google.com/',
                'https://www.googleapis.com/auth/activity',
                'https://www.googleapis.com/auth/drive.metadata.readonly',
                'https://www.googleapis.com/auth/admin.directory.group'])
    delegate_credentials = credentials.create_delegated(email)
    return delegate_credentials.authorize(httplib2.Http())


def backup_info(filename=''):
    path = os.path.expanduser('~/Documents/backup/')
    prefix = path+'backup'
    if not (filename == ''):
        prefix = path + filename
    # backupUsers
    with open(prefix + 'Users.json', 'w') as f:
        user_info = get_users_all_info()
        f.write(json.dumps(user_info, indent=4))
    print(prefix + 'Users.json done')
    # backupGroups
    with open(prefix + 'Groups.json', 'w') as f:
        group_info = get_groups_all_info()
        f.write(json.dumps(group_info, indent=4))
    print(prefix + 'Groups.json done')
    # backupUsersInGroups
    with open(prefix + 'UsersInGroups.json', 'w') as f:
        group_info = get_all_users_in_groups()
        f.write(json.dumps(group_info, indent=4))
    print(prefix + 'UsersInGroups.json done')
    # backupUsersAliases
    with open(prefix + 'UsersAliases.json', 'w') as f:
        info = get_users_aliases()
        f.write(json.dumps(info, indent=4))
    print(prefix + 'UsersAliases.json done')
    # backupGroupsAliases
    with open(prefix + 'GroupsAliases.json', 'w') as f:
        info = get_group_aliases()
        f.write(json.dumps(info, indent=4))
    print(prefix + 'GroupsAliases.json done')


def change_mail_domain(email, domain):
    if not '@' + siteName + '.' + domain in email:
        return email.split('@')[0] + '@' + siteName + '.' + domain
    return email


def create_domains_alias(email, domain, is_default=True):
    new_email = change_mail_domain(email, domain)
    send_as_resource = {"sendAsEmail": new_email,
                        "isDefault": is_default,
                        "replyToAddress": new_email,
                        "displyaName": get_fullname(email),
                        "isPrimary": False,
                        "treatAsAlias": False
                        }
    http = get_credentials(email)
    gmail = discovery.build('gmail', 'v1', http=http)
    return gmail.users().settings().sendAs().create(userId='me', body=send_as_resource).execute()


def create_newname_alias(email, alias, is_default=False):
    body = {'alias': alias}
    http = get_credentials(adminEmail)
    service = discovery.build('admin', 'directory_v1', http=http)
    try:
        response = service.users().aliases().insert(userKey=email, body=body).execute()
        return response
    except Exception as e:
        return sys.exc_info()[0], ":", e


def create_newname_groups_alias(email, alias):
    body = {'alias': alias}
    http = get_credentials(adminEmail)
    service = discovery.build('admin', 'directory_v1', http=http)
    try:
        response = service.groups().aliases().insert(groupKey=email, body=body).execute()
        return response
    except Exception as e:
        return sys.exc_info()[0], ':', e


def create_domain_alias(email):
    email_domain = change_mail_domain(email=email, domain='com')
    send_as_resource = {
        "sendAsEmail": email_domain,
        "isDefault": True,
        "replyToAddress": email_domain,
        "displayName": get_fullname(email),
        "isPrimary": False,
        "treatAsAlias": False
    }
    print(send_as_resource)
    http = get_credentials(email)
    gmail = discovery.build('gmail', 'v1', http=http)
    response = gmail.users().settings().sendAs().create(userId='me', body=send_as_resource).execute()
    print(response)


def create_signature(post, name, surname='', telephone='', skype=''):
    template = ('<div style="font-family:helvetica,arial;font-size:13px">Regards,</div>'
                '<div style="font-family:helvetica,arial;font-size:13px"><b>')
    if surname == '':
        template = template + name
    else:
        template = template + name + ' ' + surname
    template = template + '</b><div style="font-family:helvetica,arial;font-size:13px"><i>' + post + '</i>'
    if not skype == '':
        template = template + '<div style="font-family:helvetica,arial;font-size:13px">Skype: ' + skype + '</div>'
    if not telephone == '':
        template = template + '<div>Tel: <a dir="ltr" href="tel:+%207%20916%207183639" target="_blank">' + telephone + '</a></div>'
    template += '<div><img src="https://ci4.googleusercontent.com/proxy/vsSFXEeaHOwP-g7pKbqB00dUtC1wwA1I_iG1KWsVXUfKDQg6Vvasf9u2LXsT1QYZI1YruapYZg=s0-d-e1-ft#http://cdn.pixapi.net/pixonic.png" width="250" class="CToWUd"></div><div><span style="color:rgb(0,0,0);font-size:9px">This e-mail and the attachments, if any, are confidential and may be legally privileged, and are for the sole use of the intended recipient. If you are not the intended recipient of this e-mail or any part of it, please delete it from your computer.</span></div></div></div>'
    return {'signature': template}


def create_full_user(email, firstname, secondname, post, department, telephone, skype):
    print("Create user " + email)
    create_user(email=email, firstname=firstname, secondname=secondname)
    print("Wait 15s")
    time.sleep(15)
    print("Update information about user")
    update_information(userEmail=email, post=post, department=department, telephone=telephone)
    print("Create signature")
    signature = create_signature(post=post, name=firstname, surname=secondname, telephone=telephone, skype=skype)
    print('Edit signature for ru')
    edit_signature(email=email, signature=signature)
    print("wait 1 minute")
    time.sleep(60)
    print("Create alias com")
    create_domain_alias(email)
    time.sleep(15)
    print('Edit signature for com ' + change_mail_domain(email=email, domain='com')(email))
    edit_signature(email=change_mail_domain(email=email, domain='com'), signature=signature)


def create_user(email, firstname, secondname):
    http = get_credentials(adminEmail)
    service = discovery.build('admin', 'directory_v1', http=http)
    user_info = {
        'primaryEmail': email,
        'name': {'givenName': firstname, 'familyName': secondname},
        'password': 'Zz123456'}
    service.users().insert(body=user_info).execute()


def edit_signature(email, signature):
    http = get_credentials(email)
    gmail = discovery.build('gmail', 'v1', http=http)
    print(signature)
    response = gmail.users().settings().sendAs().patch(userId='me', sendAsEmail=email, body=signature).execute()
    print(response)


def get_alias_of_user(usermail):
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().aliases().list(userKey=usermail).execute()
    result = []
    try:
        for a in data['aliases']:
            result.append(a['alias'])
    except KeyError:
        pass
    return result


def get_group_aliases():
    groups = get_groups_all_info()
    root = {}
    for group in groups:
        root[group] = {}
        try:
            root[group]['aliases'] = groups[group]['aliases']
        except KeyError:
            pass
        try:
            root[group]['nonEditableAliases'] = groups[group]['nonEditableAliases']
        except KeyError:
            pass
    return root


def get_signature(email):
    http = get_credentials(email)
    gmail = discovery.build('gmail', 'v1', http=http)
    response = gmail.users().settings().sendAs().get(sendAsEmail=email, userId='me').execute()
    return response['signature']


def get_fullname(email):
    return get_user(email)['name']['fullName']


def get_post(email):
    return get_user(email)['organizations'][0]['title']


def get_users():
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


def get_users_aliases():
    users = get_users_all_info()
    root = {}
    for user in users:
        root[user] = {}
        try:
            root[user]['aliases'] = users[user]['aliases']
        except KeyError:
            pass
        try:
            root[user]['nonEditableAliases'] = users[user]['nonEditableAliases']
        except KeyError:
            pass
    return root


def get_users_all_info():
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().list(domain=mainDomain).execute()
    result = {}
    for a in data['users']:
        result[a['primaryEmail']] = a
        # result.append(a)
    while 'nextPageToken' in data:
        data = service.users().list(domain=mainDomain, pageToken=data['nextPageToken']).execute()
        for a in data['users']:
            result[a['primaryEmail']] = a
    return result


def get_groups_all_info():
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


def get_groups_members(group_name):
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.members().list(groupKey=group_name).execute()
    result = {}
    try:
        for a in data['members']:
            try:
                result[a['email']] = a
            except KeyError:
                pass
    except KeyError:
        pass
    while 'nextPageToken' in data:
        data = service.members().list(groupKey=group_name, pageToken=data['nextPageToken']).execute()
        for a in data['members']:
            result[a['email']] = a
    return result


def get_all_users_in_groups():
    result = {}
    for group_name in get_groups_all_info():
        print(group_name)
        # result[groupName['email']] = groupName['email']
        result[group_name] = get_groups_members(group_name)
    return result


def get_user(email):
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    data = service.users().get(userKey=email).execute()
    return data


def load_json_file(filename):
    with open(filename, 'r') as file:
        result = json.load(file)
    return result


def mail_to_new_domain(email):
    if not ('@' + siteName + '.com') in email:
        return email.split('@')[0] + '@' + siteName + '.com'
    return email


def move_users(domain):
    # Получаем список всех пользователей
    users = get_users()
    # Переименовываем всех пользователей, кроме adminEmail
    # print(json.dumps(users, indent=4))
    for a in users:
        if adminEmail in a[0]: continue
        try:
            rename_user(a[0], domain)
            print('Done: ' + a[0])
        except Exception as e:
            print(sys.exc_info()[0], ":", e)
            print('Failed: ' + a[0])
    print("Users done!!!")


def move_groups(domain): # domain='com'
    # Получаем список групп
    groups = get_groups_all_info()
    for group in groups:
        try:
            rename_group(groups[group]['email'], domain)
            print('Done: ' + groups[group]['email'])
        except Exception as e:
            print(sys.exc_info()[0], ":", e)
            print('Failed: ' + groups[group]['email'])
    print("Groups are done!!!")


def move_users_and_groups(domain):
    move_users(domain)
    move_groups(domain)


def rename_group(group, domain):
    patch = {'email': change_mail_domain(group, domain)}
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    service.groups().update(groupKey=group, body=patch).execute()


def rename_user(email, domain):
    patch = {'primaryEmail': change_mail_domain(email=email, domain=domain)}
    #print(json.dumps(patch, indent=4))
    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    #print(email, patch)
    service.users().update(userKey=email, body=patch).execute()


def restore_alias_groups(filename):
    groups = load_json_file(filename)
    result = {}
    for group in groups:
        try:
            aliases = groups[group]['aliases']
            if len(aliases) > 0:
                result[group] = aliases
        except:
            pass
    aliases = result
    flag = True
    for group in aliases:
        for alias in aliases[group]:
            try:
                create_newname_groups_alias(email=group, alias=change_mail_domain(email=alias, domain=newDomain))
                print("Done:", alias)
            except Exception as e:
                print("Failed:", alias)
                print(sys.exc_info()[0], ":", e)
                flag = False
    if flag:
        print("All groups' aliases have been restored")
    else:
        print("We have a problem with groups")


def restore_alias_users(filename):
    users = load_json_file(filename)
    result = {}
    for user in users:
        try:
            aliases = users[user]['aliases']
            if len(aliases) > 0:
                result[user] = aliases
        except:
            pass
    aliases = result
    flag = True
    for user in aliases:
        for alias in aliases[user]:
            try:
                create_newname_alias(email=user, alias=change_mail_domain(alias, newDomain))
                print("Done: ", alias)
            except Exception as e:
                print("Failed: ", alias)
                print(sys.exc_info()[0], ":", e)
                flag = False
    if flag:
        print("All users' aliases have been restored")
    else:
        print("We have a problem with users")


def save_output_to_file(lines, filename):
    with open(filename, 'w') as file:
        for a in lines:
            file.write(str(a))
            file.write('\n')


def update_information(userEmail, post, department, telephone):
    patch = {
        'organizations': [{'title': post, 'primary': True, 'customType': '',
                           'department': department, 'description': '', 'costCenter': ''}],
        'phones': [{'value': telephone, 'type': 'home'}]}

    http = get_credentials(adminEmail)
    service = build('admin', 'directory_v1', http=http)
    service.users().update(userKey=userEmail, body=patch).execute()


def update_signature_from_google(email): #For 'ru' & 'com'
    full_name = get_user(email=email)['name']['fullName']
    post = get_user(email=email)['organizations'][0]['title']
    signature = create_signature(name=full_name, post=post)
    #Update ru
    try:
        edit_signature(email=change_mail_domain(email=email, domain='ru'), signature=signature)
    except Exception as e:
        print(change_mail_domain(email=email, domain='ru'), '\t', sys.exc_info()[0], ":", e)
    #Update com
    try:
        edit_signature(email=change_mail_domain(email=email, domain='com'), signature=signature)
    except Exception as e:
        print(change_mail_domain(email=email, domain='com'), '\t', sys.exc_info()[0], ":", e)


def main():
    now = datetime.now()
    hFile = str(now.month) + str(now.day)
    global adminEmail # It needs to modify global copy of globvar
    
#Снять следующий комментарий перед восстановлением из бекапа
#"""  
    #Делаем backup
    print("Делаем backup")
    backup_info(siteName + hFile)
    #Делаем миграцию в новый домен
    print("Делаем миграцию в новый номен")
    move_users_and_groups(newDomain)
    #Мигририум собственный аккаунт
    print("Мигрируем собственный аккаунт")
    rename_user(adminEmail, newDomain)
    
    #Ожидаем завершения ручных операций по смене главного домена
    #Авторизуем api в новом домене
    """
    print("Переименовываем adminEmail")
    adminEmail = change_mail_domain(adminEmail, newDomain)
    print(adminEmail)
    #restore alias
    path = os.path.expanduser('~/Documents/backup/')
    usersAliases = path + siteName + hFile + "UsersAliases.json"
    print("Восстанавливаем пользовательские aliases")
    restore_alias_users(usersAliases)
    groupsAliases = path + siteName + hFile + "GroupsAliases.json"
    print("Восстанавливаем групповые aliases")
    restore_alias_groups(groupsAliases)
#"""

if __name__ == '__main__':
    main()
