from PyQt5 import QtCore, QtGui, QtWidgets
from logic import add_contact_to_mailbox , create_contact_data
from CSVimport import importAliasesFromCSV
from remove import remove_alias_from_selected_mailboxes
from dotenv import load_dotenv
import os

# Загружаем переменные окружения
load_dotenv()

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        # MainWindow.setMinimumSize(830,400)
        MainWindow.setMinimumSize(530,300)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        # Создаем главный вертикальный layout
        self.mainLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        
        # Настраиваем QTabWidget
        self.tabWidget = QtWidgets.QTabWidget()
        self.mainLayout.addWidget(self.tabWidget)

        # Вкладка "Додати"
        self.addTab = QtWidgets.QWidget()
        addTabLayout = QtWidgets.QVBoxLayout(self.addTab)
        self.addTabButton = QtWidgets.QPushButton("Додати Контакт")
        
        # Добавляем отступ сверху
        addTabLayout.addSpacing(10)
        
        # Создаем горизонтальный layout для выравнивания кнопки справа
        addButtonLayout = QtWidgets.QHBoxLayout()
        addButtonLayout.addStretch()
        addButtonLayout.addWidget(self.addTabButton)
        addTabLayout.addLayout(addButtonLayout)
        # Добавляем растягивающийся элемент снизу
        addTabLayout.addStretch()
        self.tabWidget.addTab(self.addTab, "Додати")

        # Вкладка "Видалити"
        self.removeTab = QtWidgets.QWidget()
        removeTabLayout = QtWidgets.QVBoxLayout(self.removeTab)
        self.removeTabButton = QtWidgets.QPushButton("Видалити Контакт")
        
        # Добавляем отступ сверху
        removeTabLayout.addSpacing(10)
        
        removeButtonLayout = QtWidgets.QHBoxLayout()
        removeButtonLayout.addStretch()
        removeButtonLayout.addWidget(self.removeTabButton)
        removeTabLayout.addLayout(removeButtonLayout)
        # Добавляем растягивающийся элемент снизу
        removeTabLayout.addStretch()
        self.tabWidget.addTab(self.removeTab, "Видалити")

        # Вкладка "Змінити"
        self.editTab = QtWidgets.QWidget()
        editTabLayout = QtWidgets.QVBoxLayout(self.editTab)
        self.editTabButton = QtWidgets.QPushButton("Змінити Контакт")
        
        # Добавляем отступ сверху
        editTabLayout.addSpacing(10)
        
        editButtonLayout = QtWidgets.QHBoxLayout()
        editButtonLayout.addStretch()
        editButtonLayout.addWidget(self.editTabButton)
        editTabLayout.addLayout(editButtonLayout)
        # Добавляем растягивающийся элемент снизу
        editTabLayout.addStretch()
        self.tabWidget.addTab(self.editTab, "Змінити")

        # Создаем layout для комбобоксов и меток
        comboBoxContainer = QtWidgets.QWidget()
        comboBoxMainLayout = QtWidgets.QHBoxLayout(comboBoxContainer)
        
        # Создаем вертикальные layouts для каждой пары метка+комбобокс
        mailboxLayout = QtWidgets.QVBoxLayout()
        aliasLayout = QtWidgets.QVBoxLayout()
        
        # Настраиваем метки
        self.label_mail = QtWidgets.QLabel("Поштові Скриньки")
        self.label_alias = QtWidgets.QLabel("Аліаси")
        
        # Настройка comboBox с чекбоксами
        self.comboBox = QtWidgets.QComboBox()
        self.comboBox.setView(QtWidgets.QListView())
        self.comboBox.setFixedWidth(150)
        
        # Настройка comboBox_2 для алиасов
        self.comboBox_2 = QtWidgets.QComboBox()
        self.comboBox_2.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        self.comboBox_2.setFixedWidth(150)
        
        # Инициализируем список алиасов
        self.alias_list = []
        
        # Добавляем элементы в вертикальные layouts
        mailboxLayout.addWidget(self.label_mail)
        mailboxLayout.addWidget(self.comboBox)
        
        aliasLayout.addWidget(self.label_alias)
        aliasLayout.addWidget(self.comboBox_2)
        
        # Создаем вертикальный layout для кнопок управления
        buttonLayout = QtWidgets.QVBoxLayout()
        
        # Создаем кнопки
        self.addAliasButton = QtWidgets.QPushButton("Додати Аліас")
        self.removeAliasButton = QtWidgets.QPushButton("Видалити Аліас")
        self.importCsvButton = QtWidgets.QPushButton("Імпорт з CSV")
        
        # Добавляем пустое пространство сверху кнопок для выравнивания с комбобоксами
        buttonLayout.addSpacing(self.label_mail.sizeHint().height())
        buttonLayout.addWidget(self.addAliasButton)
        buttonLayout.addWidget(self.removeAliasButton)
        buttonLayout.addWidget(self.importCsvButton)
        
        # Добавляем все layouts в основной горизонтальный
        comboBoxMainLayout.addLayout(mailboxLayout)
        comboBoxMainLayout.addLayout(aliasLayout)
        comboBoxMainLayout.addLayout(buttonLayout)
        
        # Добавляем контейнер в главный layout
        self.mainLayout.addWidget(comboBoxContainer)

        # Получение и преобразование строки в список для почтовых ящиков
        email_list_str = os.getenv("EMAIL_LIST", "")
        email_list = email_list_str.split(",")
        self.setupCheckableComboBox(self.comboBox, email_list)

        # Обновляем комбобокс алиасов
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
        # Добавляем пункт "Виберіть Аліас" всегда на первом месте
        placeholder_item = QtGui.QStandardItem("Виберіть Аліас")
        placeholder_item.setFlags(QtCore.Qt.ItemIsEnabled)  # Только для чтения
        model.appendRow(placeholder_item)

        for alias in self.alias_list:
            # Форматируем строку с двумя полями
                item_text = f"{alias['First Name']} <{alias['EmailAddress']}>"
                item = QtGui.QStandardItem(item_text)
                model.appendRow(item)
        self.comboBox_2.setCurrentText("Виберіть Аліас")

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
        if current_index <=0 :
            # Всплывающее окно предупреждения
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Помилка Видалення")
            msg_box.setText("Виберіть Аліас для видалення.")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg_box.exec_()
            return

          # Удаление алиаса: если индекс выбранного элемента больше 0, удаляем алиас
        selected_alias_index = current_index - 1 
        if selected_alias_index < len(self.alias_list):
            del self.alias_list[selected_alias_index]  # Учитываем смещение, так как индекс 0 - это "Виберіть Аліас"
            self.updateAliasComboBox()  # Обновляем список
        else:
            # Если индекс выходит за пределы списка
            msg_box = QtWidgets.QMessageBox()
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Помилка Видалення")
            msg_box.setText("Не вдалося знайти аліас для видалення.")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg_box.exec_()
            return

        
    
    # Функция для обработки выбранного алиаса из комбобокса
    def get_selected_alias_from_combobox(self,comboBox):
        """Получить данные выбранного алиаса из combobox."""
        index = comboBox.currentIndex()
        if index <= 0:  # Если выбран пункт "Виберіть Аліас" или ничего не выбрано
            return None
        adjusted_index = index - 1  # Корректируем индекс для доступа к self.alias_list
        if 0 <= adjusted_index < len(self.alias_list):
            return self.alias_list[adjusted_index]
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
    
    # Включаем поддержку масштабирования для дисплеев с высоким разрешением
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
