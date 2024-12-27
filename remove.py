from PyQt5 import QtWidgets,QtCore
from logic import remove_alias_from_mailbox

def remove_alias_from_selected_mailboxes(self): # Получаем выбранный алиас
    selected_alias = self.get_selected_alias_from_combobox(self.comboBox_2)
    if not selected_alias:
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setWindowTitle("Помилка")
        msg_box.setText("Будь ласка, виберіть аліас для видалення.")
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

    # Удаляем алиас из каждого выбранного ящика
    errors = []
    for mailbox in selected_mailboxes:
        try:
            remove_alias_from_mailbox(mailbox, selected_alias["EmailAddress"])
        except Exception as e:
            errors.append(f"Не вдалось видалити {selected_alias['EmailAddress']} з {mailbox}: {e}")

    # Итоговое сообщение
    if not errors:
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Information)
        msg_box.setWindowTitle("Успіх")
        msg_box.setText(f"Аліас {selected_alias['EmailAddress']} успішно видалений з вибраних скриньок.")
        msg_box.exec_()
    else:
        error_msg = "\n".join(errors)
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setWindowTitle("Помилки")
        msg_box.setText(f"Деякі помилки при видаленні:\n{error_msg}")
        msg_box.exec_()