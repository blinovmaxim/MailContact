from PyQt5 import QtWidgets, QtCore
from logic import get_access_token
import requests

def edit_action(self):
    # Получаем выбранный алиас
    selected_alias = self.get_selected_alias_from_combobox(self.comboBox_2)
    if not selected_alias:
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setWindowTitle("Помилка")
        msg_box.setText("Будь ласка, виберіть аліас для редагування.")
        msg_box.exec_()
        return

    # Получаем выбранные почтовые ящики
    selected_mailboxes = []
    for index in range(1, self.comboBox.count()):
        item = self.comboBox.model().item(index)
        if item.checkState() == QtCore.Qt.Checked:
            selected_mailboxes.append(item.text())

    if not selected_mailboxes:
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setWindowTitle("Помилка")
        msg_box.setText("Будь ласка, виберіть хоча б одну поштову скриньку.")
        msg_box.exec_()
        return

    # Диалог для ввода новых данных
    new_name, ok1 = QtWidgets.QInputDialog.getText(
        None, 
        "Редагування контакту", 
        "Нове ім'я:",
        text=selected_alias["First Name"]
    )
    if not ok1:
        return

    new_email, ok2 = QtWidgets.QInputDialog.getText(
        None, 
        "Редагування контакту", 
        "Нова електронна адреса:",
        text=selected_alias["EmailAddress"]
    )
    if not ok2:
        return

    # Добавляем диалог для мобильного телефона
    current_mobile = selected_alias.get("MobilePhone", "")
    new_mobile_phone, ok4 = QtWidgets.QInputDialog.getText(
        None,
        "Редагування контакту",
        "Новий мобільний телефон:",
        text=current_mobile
    )
    if not ok4:
        return

    # Обновляем контакт в каждом выбранном ящике
    errors = []
    success = []
    
    for mailbox in selected_mailboxes:
        try:
            update_contact(
                mailbox, 
                selected_alias["EmailAddress"], 
                new_name, 
                new_email,
                new_mobile_phone.strip()
            )
            success.append(mailbox)
        except Exception as e:
            errors.append(f"Помилка в {mailbox}: {str(e)}")

    # Обновляем локальный список алиасов
    for alias in self.alias_list:
        if (alias["First Name"] == selected_alias["First Name"] and 
            alias["EmailAddress"] == selected_alias["EmailAddress"]):
            alias["First Name"] = new_name
            alias["EmailAddress"] = new_email
            alias["MobilePhone"] = new_mobile_phone.strip()
            break

    # Обновляем комбобокс
    self.updateAliasComboBox()

    # Показываем результат
    result_message = ""
    if success:
        result_message += f"Успішно оновлено в:\n{', '.join(success)}\n\n"
    if errors:
        result_message += f"Помилки:\n{'\n'.join(errors)}"

    msg_box = QtWidgets.QMessageBox()
    msg_box.setIcon(QtWidgets.QMessageBox.Information)
    msg_box.setWindowTitle("Результат редагування")
    msg_box.setText(result_message)
    msg_box.exec_()

def update_contact(user_id, old_email, new_name, new_email, mobile_phone=None):
    """Обновляет существующий контакт в почтовом ящике."""
    access_token = get_access_token()
    
    # Сначала найдем ID контакта
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/contacts?$filter=emailAddresses/any(e:e/address eq '{old_email}')"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        raise Exception(f"Помилка пошуку контакту: {response.status_code}")

    contacts = response.json().get('value', [])
    if not contacts:
        raise Exception(f"Контакт не знайдено")

    contact_id = contacts[0]['id']

    # Обновляем контакт
    update_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/contacts/{contact_id}"
    update_data = {
        "givenName": new_name,
        "emailAddresses": [
            {
                "address": new_email
            }
        ],
        "mobilePhone": mobile_phone if mobile_phone else ""
    }

    update_response = requests.patch(update_url, headers=headers, json=update_data)
    if update_response.status_code != 200:
        raise Exception(f"Помилка оновлення: {update_response.status_code}")
