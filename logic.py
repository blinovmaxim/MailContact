import requests
from dotenv import load_dotenv
import os
import sys

def get_base_path():
    """Получаем путь к директории с исполняемым файлом"""
    if getattr(sys, 'frozen', False):
        # Если приложение скомпилировано
        return sys._MEIPASS
    else:
        # Если запущено как скрипт
        return os.path.dirname(os.path.abspath(__file__))

# Загружаем переменные окружения
env_path = os.path.join(get_base_path(), '.env')
load_dotenv(env_path)

# Параметры для работы с Azure AD
TenantId = os.getenv("TENANT_ID")
ClientId = os.getenv("CLIENT_ID")
ClientSecret = os.getenv("CLIENT_SECRET")

# Добавьте проверку
if not all([TenantId, ClientId, ClientSecret]):
    raise Exception("Не вдалося завантажити конфігурацію з .env файлу")

# Получение токена
def get_access_token():
    url = f"https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token"
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    body = {
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
        "client_id": ClientId,
        "client_secret": ClientSecret
    }
    
    response = requests.post(url, headers=headers, data=body)
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        raise Exception("Ошибка получения токена доступа")
    

def add_contact_to_mailbox(user_id, contact_data):
    access_token = get_access_token()
    
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/contacts"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=contact_data)

    if response.status_code == 201:
        print(f"Контакт {contact_data['givenName']}  добавлен в почтовый ящик {user_id}.")
    else:
        print(f"Ошибка при добавлении контакта в ящик {user_id}: {response.status_code} - {response.text}")

# Функция для создания данных контакта
def create_contact_data(given_name, email_address, mobile_phone=None):
    contact_data = {
        "givenName": given_name,
        "emailAddresses": [
            {
                "address": email_address
            }
        ],
        "mobilePhone": mobile_phone if mobile_phone else ""
    }
    return contact_data


def get_contact_id(user_id, email_address):
    access_token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/contacts"
    headers = {
        "Authorization": f"Bearer {access_token}",
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        contacts = response.json().get('value', [])
        for contact in contacts:
            if email_address in [email['address'] for email in contact.get('emailAddresses', [])]:
                return contact['id']
        raise Exception(f"Контакт с email {email_address} не найден.")
    else:
        raise Exception(f"Ошибка получения списка контактов: {response.status_code} - {response.text}")



def remove_alias_from_mailbox(user_id, alias_email):
    """
    Удаляет алиас (контакт) из почтового ящика пользователя.
    
    :param user_id: Идентификатор почтового ящика.
    :param alias_email: Адрес алиаса, который нужно удалить.
    """
    access_token = get_access_token()
    
    # URL для поиска контактов по email
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/contacts?$filter=emailAddresses/any(e:e/address eq '{alias_email}')"
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        contacts = response.json().get('value', [])
        if not contacts:
            raise Exception(f"Контакт с email {alias_email} не найден в почтовом ящике {user_id}.")
    
        for contact in contacts:
                contact_id = contact.get('id')
                delete_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/contacts/{contact_id}"
                delete_response = requests.delete(delete_url, headers=headers)

                if delete_response.status_code == 204:
                    print(f"Контакт с email {alias_email} успешно удален из ящика {user_id}.")
                else:
                    raise Exception(f"Ошибка при удалении контакта {alias_email}: {delete_response.status_code} - {delete_response.text}")
    else:
        raise Exception(f"Ошибка при поиске контактов: {response.status_code} - {response.text}")