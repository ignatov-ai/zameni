###################################################################################
######################## ИСПОЛЬЗУЕМЫЕ ФУНКЦИИ И ПРОЦЕДУРЫ #########################
###################################################################################

def bolnichniyExcelLoad(self):
    # получение данных из файла со списком больничных листов
    load_bolnichniy_file = 'test.xlsx'
    bolnichnie_data = pd.read_excel(load_bolnichniy_file, 'Учет больничных листов')
    bolnichnie_data.fillna('', inplace=True)

    self.bolnichniy_table.setRowCount(bolnichnie_data.shape[0])
    self.bolnichniy_table.setColumnCount(bolnichnie_data.shape[1])
    self.bolnichniy_table.setHorizontalHeaderLabels(bolnichnie_data.columns)

    for row in bolnichnie_data.iterrows():
        values = row[1]
        for col_index, value in enumerate(values):
            tableItem = QTableWidgetItem(str(value))
            self.bolnichniy_table.setItem(row[0], col_index, tableItem)


def zameniExcelLoad(self):
    # получение данных из файла со списком больничных листов
    load_zameni_file = 'zameni.xlsx'
    zameni_data = pd.read_excel(load_zameni_file, 'Замены')
    zameni_data.fillna('', inplace=True)

    self.zameni_table.setRowCount(zameni_data.shape[0])
    self.zameni_table.setColumnCount(zameni_data.shape[1])
    self.zameni_table.setHorizontalHeaderLabels(zameni_data.columns)

    for row in zameni_data.iterrows():
        values = row[1]
        for col_index, value in enumerate(values):
            tableItem = QTableWidgetItem(str(value))
            self.zameni_table.setItem(row[0], col_index, tableItem)


# определение текущей вкладки
def currentTabNumber(self, index):
    self.tab_index = self.tabs.currentIndex()
    if self.tab_index == 2:
        self.bolnichniyExcelLoad()
    elif self.tab_index == 3:
        self.zameniExcelLoad()


# вывод в выпадающий список учителей по фильтру из поля поиска
def teachFind(self, text):
    fioSrez = fix_input(text.capitalize())
    lenFioSrez = len(fioSrez)
    sotrudniki_find = []
    for i in range(len(sotrudniki_fio)):
        if sotrudniki_fio[i][6:6 + lenFioSrez] == fioSrez:
            sotrudniki_find.append(sotrudniki_fio[i])
    if len(sotrudniki_find) > 0:
        self.teach_select.clear()
        self.teach_select.addItems(sotrudniki_find)
    else:
        self.teach_select.addItems(sotrudniki_fio)


def teachFind_2(self, text):
    fioSrez = fix_input(text.capitalize())
    lenFioSrez = len(fioSrez)
    sotrudniki_find_2 = []
    for i in range(len(sotrudniki_fio)):
        if sotrudniki_fio[i][6:6 + lenFioSrez] == fioSrez:
            sotrudniki_find_2.append(sotrudniki_fio[i])
    if len(sotrudniki_find_2) > 0:
        self.teach_select_2.clear()
        self.teach_select_2.addItems(sotrudniki_find_2)
    else:
        self.teach_select_2.addItems(sotrudniki_fio)


def bolnichniy_add():
    fio_text = teach_select.currentText()
    id = fio_text[:4]
    for i in range(len(sotrudniki)):
        if sotrudniki[i][3] == id:
            break
    zamena_zapis = [sotrudniki[i][j] for j in range(5)]
    zamena_zapis.append(self.zamena_dateStart.text())
    zamena_zapis.append(self.zamena_dateEnd.text())
    # print(zamena_zapis)

    wb = load_workbook(filename=bolnichniy_book_name)
    ws = wb['Учет больничных листов']
    ws.append(zamena_zapis)
    sheet = wb.active
    # выравнивание столбцов D E F G по центру
    for c in 'DEFG':
        currRow = c + str(sheet.max_row)
        cell = sheet[str(currRow)]
        alignment = Alignment(horizontal="center", vertical="center")
        cell.alignment = alignment
    wb.save(filename=bolnichniy_book_name)
    wb.close()


def zamena_date_select(self):
    selected_date = self.zamena_select_2.text()
    # print(selected_date)
    self.lbl_zamena_s_2.setText(selected_date)
    return selected_date


def zamena_teach_select(self):
    selected_fio_text = self.teach_select_2.currentText()
    # print(selected_fio_text)
    self.lbl_teach_select_2.setText(selected_fio_text)
    return selected_fio_text[6:]


def zamena_lessons_build(self, tab2):
    day, month, year = (int(x) for x in self.zamena_select_2.text().split('.'))
    sel_date = datetime.weekday(datetime(year, month, day))
    sel_teach = self.teach_select_2.currentText()[:4]

    if sel_date == 5 or sel_date == 6:
        self.lbl_zamena_s_2.setText('<p style="color: rgb(250, 55, 55);">ВЫБРАНА НЕВЕРНАЯ ДАТА</p>', )
    else:
        for teach_id in range(len(sotrudniki)):
            if raspisanie[teach_id][0] == sel_teach:
                break

        for sotr_teach_id in range(len(sotrudniki)):
            if sotrudniki[sotr_teach_id][3] == sel_teach:
                break
        # print('Индекс выбранного учителя в массиве сотрудников:', teach_id)

        teach_rasp = raspisanie[teach_id][4 + sel_date * 10: 4 + (sel_date + 1) * 10]

        if teach_rasp[0] != '-':
            self.lbl_les_1_label.setText('1.     ' + teach_rasp[0])
            self.les_zamena_add_btn_1.setEnabled(True)
        else:
            self.lbl_les_1_label.setText('1.     ' + '----------')
            self.les_zamena_add_btn_1.setEnabled(False)

        if teach_rasp[1] != '-':
            self.lbl_les_2_label.setText('2.     ' + teach_rasp[1])
            self.les_zamena_add_btn_2.setEnabled(True)
        else:
            self.lbl_les_2_label.setText('2.     ' + '----------')
            self.les_zamena_add_btn_2.setEnabled(False)

        if teach_rasp[2] != '-':
            self.lbl_les_3_label.setText('3.     ' + teach_rasp[2])
            self.les_zamena_add_btn_3.setEnabled(True)
        else:
            self.lbl_les_3_label.setText('3.     ' + '----------')
            self.les_zamena_add_btn_3.setEnabled(False)

        if teach_rasp[3] != '-':
            self.lbl_les_4_label.setText('4.     ' + teach_rasp[3])
            self.les_zamena_add_btn_4.setEnabled(True)
        else:
            self.lbl_les_4_label.setText('4.     ' + '----------')
            self.les_zamena_add_btn_4.setEnabled(False)

        if teach_rasp[4] != '-':
            self.lbl_les_5_label.setText('5.     ' + teach_rasp[4])
            self.les_zamena_add_btn_5.setEnabled(True)
        else:
            self.lbl_les_5_label.setText('5.     ' + '----------')
            self.les_zamena_add_btn_5.setEnabled(False)

        if teach_rasp[5] != '-':
            self.lbl_les_6_label.setText('6.     ' + teach_rasp[5])
            self.les_zamena_add_btn_6.setEnabled(True)
        else:
            self.lbl_les_6_label.setText('6.     ' + '----------')
            self.les_zamena_add_btn_6.setEnabled(False)

        if teach_rasp[6] != '-':
            self.lbl_les_7_label.setText('7.     ' + teach_rasp[6])
            self.les_zamena_add_btn_7.setEnabled(True)
        else:
            self.lbl_les_7_label.setText('7.     ' + '----------')
            self.les_zamena_add_btn_7.setEnabled(False)

        if teach_rasp[7] != '-':
            self.lbl_les_8_label.setText('8.     ' + teach_rasp[7])
            self.les_zamena_add_btn_8.setEnabled(True)
        else:
            self.lbl_les_8_label.setText('8.     ' + '----------')
            self.les_zamena_add_btn_8.setEnabled(False)

        if teach_rasp[8] != '-':
            self.lbl_les_9_label.setText('9.     ' + teach_rasp[8])
            self.les_zamena_add_btn_9.setEnabled(True)
        else:
            self.lbl_les_9_label.setText('9.     ' + '----------')
            self.les_zamena_add_btn_9.setEnabled(False)

        if teach_rasp[9] != '-':
            self.lbl_les_10_label.setText('10.   ' + teach_rasp[9])
            self.les_zamena_add_btn_10.setEnabled(True)
        else:
            self.lbl_les_10_label.setText('10.   ' + '----------')
            self.les_zamena_add_btn_10.setEnabled(False)


def zamena_add_form(self, i):
    num_les = i
    day, month, year = (int(x) for x in self.zamena_select_2.text().split('.'))
    date = datetime(year, month, day)
    sel_date = datetime.weekday(datetime(year, month, day))
    sel_teach_all = self.teach_select_2.currentText()
    sel_teach = self.teach_select_2.currentText()[:4]

    if sel_date == 5 or sel_date == 6:
        self.lbl_zamena_s_2.setText('<p style="color: rgb(250, 55, 55);">ВЫБРАНА НЕВЕРНАЯ ДАТА</p>', )
    else:
        for teach_id in range(len(sotrudniki)):
            if raspisanie[teach_id][0] == sel_teach:
                break

        for sotr_teach_id in range(len(sotrudniki)):
            if sotrudniki[sotr_teach_id][3] == sel_teach:
                break
        # print('Индекс выбранного учителя в массиве сотрудников:', teach_id)

        # вывод заменяемого учителя и расписание его уроков на этот день
        sel_teach_num_les = []
        for i in range(4 + sel_date * 10, 4 + (sel_date + 1) * 10):
            if raspisanie[teach_id][i] != '-':
                # сохранение в массив sel_teach_num_les уроков для замены (окна пропущены)
                sel_teach_num_les.append(i)
        # print('Номера уроков для замены:', *sel_teach_num_les)
        self.sel_predmet = raspisanie[teach_id][2]
        # вывод учителей ТОГО ЖЕ предмета
        kandidati = []
        self.send_teach_id = teach_id

        for j in range(len(raspisanie)):
            for i in sel_teach_num_les:
                if raspisanie[j][2] == self.sel_predmet and raspisanie[j][0] != sel_teach \
                        and raspisanie[j][i] == '-':
                    kand_temp = str(i - 3 - sel_date * 10) + ';' + str(raspisanie[j][0]) + ';' \
                                + str(raspisanie[j][1]) + ';' + str(raspisanie[teach_id][i])
                    kandidati.append(kand_temp)

        # вывод учителей ОСТАЛЬНЫХ предметов
        for j in range(len(raspisanie)):
            for i in sel_teach_num_les:
                if raspisanie[j][2] != self.sel_predmet and raspisanie[j][i] == '-':
                    kand_temp = str(i - 3 - sel_date * 10) + ';' + str(raspisanie[j][0]) + ';' + \
                                str(raspisanie[j][1]) + ';' + str(raspisanie[teach_id][i])
                    kandidati.append(kand_temp)

        kandidati = sorted(kandidati, key=lambda row: row[0])
        n_lessons = 10 - raspisanie[teach_id][4 + sel_date * 10: 4 + (sel_date + 1) * 10].count('-')
        # print('Количество заменяемых уроков:', n_lessons)
        # print('Количество кандидатов для замены ВСЕХ уроков:', len(kandidati))

        self.win_zamena = zamena_add_window(teach_id, num_les, sel_teach_all, kandidati,
                                            date, sel_date, self.sel_predmet)
        self.win_zamena.show()