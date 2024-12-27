from PyQt5 import QtCore, QtGui, QtWidgets
from logic import add_contact_to_mailbox , create_contact_data
from CSVimport import importAliasesFromCSV
from remove import remove_alias_from_selected_mailboxes

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(530,200)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        # Создаем QTabWidget для вкладок
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 10, 532, 200))
        self.tabWidget.setObjectName("tabWidget")


        # Вкладка "Додати"
        self.addTab = QtWidgets.QWidget()
        self.addTab.setObjectName("addTab")
        self.addTabButton = QtWidgets.QPushButton("Додати Контакт", self.addTab)
        self.addTabButton.setGeometry(QtCore.QRect(360, 20, 125, 23))
        self.addTabButton.setObjectName("add_contact")
        self.tabWidget.addTab(self.addTab, "Додати")
        


        # Вкладка "Видалити"
        self.removeTab = QtWidgets.QWidget()
        self.removeTabButton = QtWidgets.QPushButton("Видалити Контакт", self.removeTab)
        self.removeTabButton.setGeometry(QtCore.QRect(360, 20, 125, 23))
        self.tabWidget.addTab(self.removeTab, "Видалити")

        # Вкладка "Змінити"
        self.editTab = QtWidgets.QWidget()
        self.editTab.setObjectName("editTab")
        self.editTabButton = QtWidgets.QPushButton("Змінити Контакт", self.editTab)
        self.editTabButton.setGeometry(QtCore.QRect(360, 20, 125, 23))
        self.tabWidget.addTab(self.editTab, "Змінити")




        self.horizontalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(0, 46, 321, 41))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(10, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_mail = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_mail.setObjectName("label_mail")
        self.horizontalLayout.addWidget(self.label_mail)

        self.label_alias = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_alias.setObjectName("label_alias")
        self.horizontalLayout.addWidget(self.label_alias)

        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(0, 70, 321, 80))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(10, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")

        # Настройка comboBox с чекбоксами
        self.comboBox = QtWidgets.QComboBox(self.horizontalLayoutWidget_2)
        self.comboBox.setView(QtWidgets.QListView())  # Используем QListView для отображения чекбоксов
        self.comboBox.setObjectName("comboBox")
        self.horizontalLayout_2.addWidget(self.comboBox)
        self.comboBox.setFixedWidth(150)

        # Список ящиков
        email_list = [
            "Виберіть поштову скриньку",
            "Всі",
            "mukachevo2@transkarpatianriteil.com.ua",
            "user2@example.com",
            "user3@example.com",
            "user4@example.com",
            "user5@example.com",
        ]

        self.setupCheckableComboBox(self.comboBox, email_list)

        # Настройка comboBox_2 для алиасов
        self.comboBox_2 = QtWidgets.QComboBox(self.horizontalLayoutWidget_2)
        self.comboBox_2.setObjectName("comboBox_2")
        self.comboBox_2.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        self.horizontalLayout_2.addWidget(self.comboBox_2)
        self.comboBox_2.setFixedWidth(150)
        

        # Кнопки для управления алиасами
        self.addAliasButton = QtWidgets.QPushButton(self.centralwidget)
        self.addAliasButton.setGeometry(QtCore.QRect(330, 100, 85, 23))
        self.addAliasButton.setObjectName("addAliasButton")

        self.removeAliasButton = QtWidgets.QPushButton(self.centralwidget)
        self.removeAliasButton.setGeometry(QtCore.QRect(430, 100, 90, 23))
        self.removeAliasButton.setObjectName("removeAliasButton")

        # Кнопка импорта алиасов из CSV
        self.importCsvButton = QtWidgets.QPushButton(self.centralwidget)
        self.importCsvButton.setGeometry(QtCore.QRect(430, 130, 90, 23))
        self.importCsvButton.setObjectName("importCsvButton")

        # Список алиасов
        self.alias_list = []
        self.updateAliasComboBox()

        

        # Подключение кнопок
        self.addTabButton.clicked.connect(self.add_contact_to_selected_mailboxes)
        self.addAliasButton.clicked.connect(self.addAlias)
        self.removeAliasButton.clicked.connect(self.removeAlias)
        self.importCsvButton.clicked.connect(self.importAliasesFromCSV)
        self.removeTabButton.clicked.connect(self.remove_alias_from_selected_mailboxes)
        self.editTabButton.clicked.connect(self.edit_action)

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    
    # Настроить QComboBox Поштові Скриньки с чекбоксами"""
    def setupCheckableComboBox(self, comboBox, items):

        model = QtGui.QStandardItemModel()
        comboBox.setModel(model)

        for i, text in enumerate(items):
            item = QtGui.QStandardItem(text)
            if i == 0:  # Первый элемент ("Виберіть поштову скриньку")
                item.setFlags(QtCore.Qt.ItemIsEnabled)  # Только для чтения
            else:
                item.setFlags(QtCore.Qt.ItemIsEnabled | QtCore.Qt.ItemIsUserCheckable)
                item.setData(QtCore.Qt.Unchecked, QtCore.Qt.CheckStateRole)
            model.appendRow(item)

        # Скрыть первый элемент ("Виберіть поштову скриньку") в выпадающем списке
        comboBox.view().setRowHidden(0, True)

        # Обработчик изменений состояния чекбоксов
        model.itemChanged.connect(lambda item: self.on_item_changed(item, model))

    # Обработчик изменения состояния чекбоксов"""
    def on_item_changed(self, item, model):
        if item.text() == "Виберіть поштову скриньку":
            return  # Игнорируем изменения для первого элемента
        model.blockSignals(True)
        if item.text() == "Всі":
            state = item.checkState()
            for row in range(1, model.rowCount()):
                model.item(row).setCheckState(state)
        else:
            all_checked = all(
                model.item(row).checkState() == QtCore.Qt.Checked
                for row in range(1, model.rowCount())
            )
            model.item(1).setCheckState(QtCore.Qt.Checked if all_checked else QtCore.Qt.Unchecked)
        model.blockSignals(False)

    # Обновить содержимое comboBox_2 Аліаси на основе alias_list"""
    def updateAliasComboBox(self):
        
        model = QtGui.QStandardItemModel()
        self.comboBox_2.setModel(model)
        self.comboBox_2.setCurrentText("Виберіть Аліас")

        if not self.alias_list:  # Если список алиасов пустой
            # Если алиасов нет, добавляем элемент "Пусто"
            item = QtGui.QStandardItem("Виберіть Аліас")
 
            model.appendRow(item)
        else:
            for alias in self.alias_list:
            # Форматируем строку с двумя полями
                item_text = f"{alias['First Name']} <{alias['EmailAddress']}>"
                item = QtGui.QStandardItem(item_text)
                model.appendRow(item)

    # Добавить новый алиас в список"""
    def addAlias(self):
        alias_name, ok = QtWidgets.QInputDialog.getText(None, "Додати аліас", "Введіть ім'я аліасу:")
        alias_email, ok_email = QtWidgets.QInputDialog.getText(None, "Додати аліас", "Введіть пошту аліасу:")

        if ok and alias_name.strip() and ok_email and alias_email.strip():
            self.alias_list.append({"First Name": alias_name.strip(), "EmailAddress": alias_email.strip()})
            self.updateAliasComboBox()

    # Удалить выбранный алиас из списка"""
    def removeAlias(self):
        if not self.alias_list:  # Проверяем, есть ли алиасы
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Помилка Видалення")
            msg_box.setText("Список Аліасов порожній.")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg_box.exec_()
            return
        current_index = self.comboBox_2.currentIndex()
        if current_index < 0 or self.comboBox_2.currentText() == "Виберіть Аліас":
            # Всплывающее окно предупреждения
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Помилка Видалення")
            msg_box.setText("Виберіть Аліас для видалення.")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg_box.exec_()
            return

        del self.alias_list[current_index]
        self.updateAliasComboBox()
        # Если список стал пустым, отображаем "Виберіть Аліас"
        if not self.alias_list:
            self.comboBox_2.setCurrentText("Виберіть Аліас")
    # Функция для обработки выбранного алиаса из комбобокса
    def get_selected_alias_from_combobox(self,comboBox):
        """Получить данные выбранного алиаса из combobox."""
        index = comboBox.currentIndex()
        if index > 0:  # Позиция 0 - это "Выберите алиас"
            selected_alias = self.alias_list[index - 1]  # Считываем алиас из списка
            return selected_alias
        else:
            return None

    # Добавить контакт в выбранные почтовые ящики"""
    def add_contact_to_selected_mailboxes(self):
        # Получаем выбранный алиас
        all_aliases = self.alias_list

        if not all_aliases:
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Помилка")
            msg_box.setText("Будь ласка, додайте хоча б один аліас.")
            msg_box.exec_()
            return

    
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
            msg_box.resize(400, 200)  # Увеличиваем размер окна
            msg_box.exec_()
            return
            # Список для успешных добавлений
        added_contacts = []
        errors = []

         # Создаем и добавляем контакт для каждого алиаса
        for alias in all_aliases:
            contact_data = create_contact_data(alias["First Name"], alias["EmailAddress"])
            # Добавляем контакт в каждый выбранный ящик
            for mailbox in selected_mailboxes:
                try:
                    add_contact_to_mailbox(mailbox, contact_data)
                    added_contacts.append(f"{contact_data['givenName']} -> {mailbox}")
                except Exception as e:
                    errors.append(f"Не вдалось додати {contact_data['givenName']} у {mailbox}: {e}")
            # Итоговое сообщение об успешных добавлениях
        if added_contacts:
            success_msg = "Додані контакти:\n" + "\n".join(added_contacts)
        else:
            success_msg = "Жоден контакт не було додано."

            # Итоговое сообщение об ошибках (если есть)
        if errors:
            error_msg = "Помилки:\n" + "\n".join(errors)
        else:
            error_msg = ""  
                    # try:
                    #     result = add_contact_to_mailbox(mailbox, contact_data)
                    #     msg_box = QtWidgets.QMessageBox()
                    #     msg_box.setIcon(QtWidgets.QMessageBox.Information)
                    #     msg_box.setWindowTitle("Успіх")
                    #     msg_box.setText(result)
                    #     msg_box.setText(f"{contact_data['givenName']} доданий у скриньку {mailbox}")
                    #     msg_box.resize(400, 200)  # Увеличиваем размер окна

                    #     msg_box.exec_()
                    # except Exception as e:
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Information)
        msg_box.setWindowTitle("Результат")
        msg_box.setText(success_msg + "\n\n" + error_msg)
        msg_box.resize(500, 300)
        msg_box.exec_()

    def importAliasesFromCSV(self):
        importAliasesFromCSV(self)

    def remove_alias_from_selected_mailboxes(self):
        remove_alias_from_selected_mailboxes(self)

    def edit_action(self):
        edit_action(self)
    
    
    
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Контакти ТКР"))
        self.addTabButton.setText(_translate("MainWindow", "Додати у Контакти"))
        self.label_mail.setText(_translate("MainWindow", "Поштові Скриньки"))
        self.label_alias.setText(_translate("MainWindow", "Аліаси"))
        self.addAliasButton.setText(_translate("MainWindow", "Додати Аліас"))
        self.removeAliasButton.setText(_translate("MainWindow", "Видалити Аліас"))
        self.importCsvButton.setText(_translate("MainWindow", "Імпорт з CSV"))
        self.removeTabButton.setText(_translate("MainWindow", "Видалити Контакти"))
        self.editTabButton.setText(_translate("MainWindow", "Змінити Контакти"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
