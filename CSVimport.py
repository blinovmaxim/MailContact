from PyQt5 import QtWidgets
import csv

def importAliasesFromCSV(self):
        """
        Открывает диалоговое окно для выбора CSV файла и загружает алиасы в comboBox_2.
        """
        options = QtWidgets.QFileDialog.Options()
        filePath, _ = QtWidgets.QFileDialog.getOpenFileName(None, "Виберіть файл CSV", "", "CSV Files (*.csv);;All Files (*)", options=options)

        if not filePath:
            return  # Если файл не выбран, ничего не делаем

        try:
            with open(filePath, mode='r', encoding='utf-8') as file:
                csv_reader = csv.DictReader(file)
                aliases = [{'First Name': row['First Name'], 'EmailAddress': row['EmailAddress']} for row in csv_reader]

            # Добавляем алиасы в список
            self.alias_list.extend(aliases)
            self.updateAliasComboBox()

            # Показываем сообщение об успешной загрузке
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Information)
            msg_box.setWindowTitle("Успіх")
            msg_box.setText(f"Завантажено {len(aliases)} аліасів.")
            msg_box.exec_()

        except Exception as e:
            # Если произошла ошибка при чтении CSV
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Помилка")
            msg_box.setText(f"Помилка при імпорті: {str(e)}")
            msg_box.exec_()