import datetime
import sys

from pyQt5.QtCore import *
from pyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from xlrd import *
from xlsxwriter import *

ui, _ = loadUiType('Library.ui')
login, _ = loadUiType('Login.ui')


class LoginUser(QWidget, login):
    def __init__(self):
        QWidget.__init__(self)
        self.setupUi(self)

        self.loginbutton()


    def loginbutton(self):
            self.pushButton.clicked.connect(self.log_in)


    def log_in(self):

            self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
            self.cur = self.db.cursor()

            username = self.lineEdit.text()
            password = self.lineEdit_2.text()

            sql = '''SELECT * FROM users'''

            self.cur.execute(sql)
            data = self.cur.fetchall()
            for row in data:
                if username == row[1] and password == row[2]:

                    self.window2 = MainApp()
                    self.window2.show()
                    self.close()



                else:
                    self.label_2.setText('**Make sure you enter correct password')

                    self.lineEdit.setText('')
                    self.lineEdit_2.setText('')














class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_Ui_Changes()
        self.Handle_Buttons()

        self.Show_Category()
        self.Show_Author()
        self.Show_Publisher()

        self.Theme_Orange()

        self.Combo_show_Category()
        self.Combo_show_Author()
        self.Combo_show_Publisher()


        #########to show data in grid table ##########

        self.Show_All_Client()
        self.Show_All_books()
        self.Show_All_Dattoday()

    def Handle_Ui_Changes(self):
        self.Hiding_Themes()
        self.tabWidget.tabBar().setVisible(False)




    def Handle_Buttons(self):
        self.pushButton_5.clicked.connect(self.Show_Themes)
        self.pushButton_21.clicked.connect(self.Hiding_Themes)

        self.pushButton.clicked.connect(self.Open_Day_To_Day)
        self.pushButton_2.clicked.connect(self.Open_Book_Tab)
        self.pushButton_3.clicked.connect(self.Open_Users_tab)
        self.pushButton_4.clicked.connect(self.Open_Setting_Tab)
        self.pushButton_30.clicked.connect(self.Open_Clients_tab)

        self.pushButton_7.clicked.connect(self.Book_Add_New)

        self.pushButton_14.clicked.connect(self.Setting_Add_Categories)
        self.pushButton_16.clicked.connect(self.Setting_Add_Publisher)
        self.pushButton_15.clicked.connect(self.Setting_Add_Author)

        self.pushButton_11.clicked.connect(self.User_Add_New)
        self.pushButton_13.clicked.connect(self.User_Login)
        self.pushButton_12.clicked.connect(self.User_Edit)

        self.pushButton_10.clicked.connect(self.Book_Search)
        self.pushButton_8.clicked.connect(self.Book_Edit)
        self.pushButton_9.clicked.connect(self.Book_Delete)

        self.pushButton_22.clicked.connect(self.Add_Client)
        self.pushButton_25.clicked.connect(self.Search_client)
        self.pushButton_23.clicked.connect(self.Edit_Client)
        self.pushButton_24.clicked.connect(self.Delete_Client)

        self.pushButton_26.clicked.connect(self.Delete_Category)
        self.pushButton_27.clicked.connect(self.Delete_Author)
        self.pushButton_39.clicked.connect(self.Delete_Publisher)


        self.pushButton_6.clicked.connect(self.Dayto_Day_Add)


        self.pushButton_28.clicked.connect(self.log_out)

        ################################################

        self.pushButton_29.clicked.connect(self.Export_Dailytdata)
        self.pushButton_31.clicked.connect(self.Export_Bookdata)
        self.pushButton_32.clicked.connect(self.Export_Clientdata)

        self.pushButton_33.clicked.connect(self.Export_User_data)

        self.pushButton_34.clicked.connect(self.Admin_log)

        ###################################################################

        self.pushButton_17.clicked.connect(self.Theme_Orange)
        self.pushButton_18.clicked.connect(self.Theme_Blue)
        self.pushButton_19.clicked.connect(self.Theme_Green)
        self.pushButton_20.clicked.connect(self.Theme_Q)

    ###################################################################


    def Show_Themes(self):
        self.groupBox.show()

    def Hiding_Themes(self):
        self.groupBox.hide()

    ##############logout##############

    def log_out(self):


        warning =  QMessageBox.warning(self,'LOGOUT','Are you sure want to logout of the application', QMessageBox.Yes | QMessageBox.No)

        if warning == QMessageBox.Yes :
            self.window1 = LoginUser()
            self.window1.show()
            self.close()






    #######################################
    ############# OPENING TAB ############

    def Open_Day_To_Day(self):
        self.tabWidget.setCurrentIndex(0)

    def Open_Book_Tab(self):
        self.tabWidget.setCurrentIndex(1)

    def Open_Users_tab(self):
        self.tabWidget.setCurrentIndex(2)

    def Open_Clients_tab(self):
        self.tabWidget.setCurrentIndex(4)

    def Open_Setting_Tab(self):
        self.tabWidget.setCurrentIndex(3)

    #######################################
    ############# Books ############

    def Book_Add_New(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        title = self.lineEdit_2.text()
        code = self.lineEdit_4.text()
        category = self.comboBox_3.currentText()
        author = self.comboBox_4.currentText()
        publisher = self.comboBox_5.currentText()
        price = self.lineEdit_3.text()
        desc = self.textEdit_2.toPlainText()

        self.cur.execute('''
        INSERT INTO book (book_name,book_description,book_code,book_category,book_author,book_publisher,book_price)VALUES (%s,%s,%s,%s,%s,%s,%s)
        ''', (title, desc, code, category, author, publisher, price))

        self.db.commit()

        self.statusBar().showMessage('              New Book Added',4000)

        self.Show_All_books()

        self.lineEdit_2.setText('')
        self.lineEdit_4.setText('')
        self.textEdit_2.setPlainText('')
        self.comboBox_3.setCurrentIndex(0)
        self.comboBox_4.setCurrentIndex(0)
        self.comboBox_5.setCurrentIndex(0)



    def Book_Search(self):

        self.lineEdit_9.setText('')
        self.lineEdit_7.setText('')
        self.comboBox_9.setCurrentIndex(-1)
        self.comboBox_10.setCurrentIndex(-1)
        self.comboBox_11.setCurrentIndex(-1)
        self.lineEdit_8.setText('')
        self.textEdit_3.setPlainText('')

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        book_search = self.lineEdit_10.text()

        sql = '''SELECT * FROM book WHERE book_name=%s'''

        self.cur.execute(sql, ([book_search]))

        data = self.cur.fetchone()

        if (data):
            self.statusBar().showMessage('              BOOK Available,4000')

            self.lineEdit_9.setText(data[1])
            self.textEdit_3.setPlainText(data[2])
            self.lineEdit_7.setText(data[3])
            self.comboBox_9.setCurrentText(data[4])
            self.comboBox_10.setCurrentText(data[5])
            self.comboBox_11.setCurrentText(data[6])
            self.lineEdit_8.setText(str(data[7]))

        else:
            self.statusBar().showMessage('              BOOK  not Available', 4000)

        self.db.commit()


    def Book_Edit(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        title = self.lineEdit_9.text()
        code = self.lineEdit_7.text()
        category = self.comboBox_9.currentText()
        author = self.comboBox_10.currentText()
        publisher = self.comboBox_11.currentText()
        price = self.lineEdit_8.text()
        desc = self.textEdit_3.toPlainText()

        search_book_t = self.lineEdit_10.text()

        self.cur.execute('''
        UPDATE book SET book_name=%s, book_description=%s, book_code=%s, book_category=%s, book_author=%s, book_publisher=%s, book_price=%s 
        WHERE book_name=%s
        ''', (title, desc, code, category, author, publisher, price, search_book_t))

        self.db.commit()

        self.statusBar().showMessage('              Book Updated', 4000)

        self.Show_All_books()

    def Book_Delete(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        title = self.lineEdit_9.text()
        code = self.lineEdit_7.text()
        category = self.comboBox_9.currentIndex()
        author = self.comboBox_10.currentIndex()
        publisher = self.comboBox_11.currentIndex()
        price = self.lineEdit_8.text()
        desc = self.textEdit_3.toPlainText()

        search_book_t = self.lineEdit_10.text()

        warning = QMessageBox.warning(self, 'Delete Book', 'Are you sure you want to delete this book',
                                      QMessageBox.Yes | QMessageBox.No)

        if warning == QMessageBox.Yes:
            sql = '''DELETE FROM book WHERE book_name=%s'''

            self.cur.execute(sql, ([search_book_t]))

            self.db.commit()
            self.statusBar().showMessage('              Book Deleted', 4000)

            self.Show_All_books()

    def Show_All_books(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute(
            ''' SELECT book_name, book_description, book_code, book_category, book_author, book_publisher, book_price FROM book ''')

        data = self.cur.fetchall()

        self.tableWidget_5.setRowCount(0)
        self.tableWidget_5.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_5.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget_5.rowCount()
            self.tableWidget_5.insertRow(row_position)

    ##################################################
    ################## USERS #########################

    def User_Add_New(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        username = self.lineEdit_11.text()
        email = self.lineEdit_12.text()
        password = self.lineEdit_13.text()
        password_again = self.lineEdit_14.text()

        if password == password_again:

            self.cur.execute('''
            INSERT INTO users (user_name,user_password,user_email) VALUES (%s,%s,%s)
            ''', (username, password, email))

            self.db.commit()
            self.statusBar().showMessage("              Added SUCCESSfully", 4000)

            self.lineEdit_11.setText('')
            self.lineEdit_12.setText('')
            self.lineEdit_13.setText('')
            self.lineEdit_14.setText('')



        else:
            self.label_11.setText('Please add valid password')

    def User_Login(self):

        username = self.lineEdit_15.text()
        password = self.lineEdit_16.text()

        sql = '''SELECT * FROM users'''

        self.cur.execute(sql)
        data = self.cur.fetchall()
        for row in data:
            if username == row[1] and password == row[2]:
                print('user match')
                self.statusBar().showMessage('              User Logged in', 4000)
                self.groupBox_4.setEnabled(True)

                self.lineEdit_17.setText(row[1])
                self.lineEdit_18.setText(row[2])
                self.lineEdit_20.setText(row[3])


            else:
                self.statusBar().showMessage('entered wrong username or password', 4000)

    def User_Edit(self):

        usered = self.lineEdit_15.text()
        usern = self.lineEdit_17.text()
        email = self.lineEdit_20.text()
        passw = self.lineEdit_18.text()
        passagain = self.lineEdit_19.text()

        if passagain == self.lineEdit_18.text():

            self.cur.execute('''
            UPDATE users SET  user_name=%s,user_password=%s,user_email=%s WHERE user_name =%s
            ''', (usern, passw, email, usered))

            self.db.commit()
            print('edited')
            self.statusBar().showMessage('              user data updated successfully', 4000)
            self.groupBox_4.setEnabled(False)

            self.lineEdit_17.setText('')
            self.lineEdit_18.setText('')
            self.lineEdit_19.setText('')
            self.lineEdit_20.setText('')
            self.lineEdit_15.setText('')
            self.lineEdit_16.setText('')

        else:
            self.statusBar().showMessage('              password not matched', 4000)

    ##################################################
    ################## SETTINGS #########################

    def Setting_Add_Categories(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        category_name = self.lineEdit_21.text()

        self.cur.execute('''
        INSERT INTO category (category_name) VALUES (%s)
        ''', (category_name,))

        self.db.commit()
        self.statusBar().showMessage("                New Category Added", 4000)
        self.lineEdit_21.setText('')
        self.Show_Category()


    def Show_Category(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT category_name FROM category''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                column = +1

            row_position = self.tableWidget_2.rowCount()
            self.tableWidget_2.insertRow(row_position)

    def Setting_Add_Publisher(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        publisher_name = self.lineEdit_23.text()

        self.cur.execute('''
            INSERT INTO publisher (publisher_name) VALUES (%s)
             ''', (publisher_name,))

        self.db.commit()
        self.statusBar().showMessage("                New Publisher Added")
        self.lineEdit_23.setText('')
        self.Show_Publisher()
        self.Combo_show_Publisher()

    def Show_Publisher(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT publisher_name FROM publisher''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                column = +1

                row_position = self.tableWidget_4.rowCount()
                self.tableWidget_4.insertRow(row_position)

    def Setting_Add_Author(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        author_nam = self.lineEdit_22.text()

        self.cur.execute('''
                   INSERT INTO author (author_name) VALUES (%s)
                   ''', (author_nam,))

        self.db.commit()
        self.statusBar().showMessage("              New Author Added", 4000)
        self.lineEdit_22.setText('')
        self.Show_Author()
        self.Combo_show_Author()

    def Show_Author(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT author_name FROM author''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
                column = +1

            row_position = self.tableWidget_3.rowCount()
            self.tableWidget_3.insertRow(row_position)

    def Delete_Author(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        del_auth = self.lineEdit_29.text()

        sql = '''DELETE FROM author WHERE author_name = %s'''

        self.cur.execute(sql ,([del_auth]))

        self.statusBar().showMessage('author removed', 3000)
        self.db.commit()
        self.Show_Author()

    def Delete_Publisher(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        del_pub = self.lineEdit_42.text()

        sql = '''DELETE FROM publisher WHERE publisher_name = %s'''

        signal = self.cur.execute(sql, ([del_pub]))
        self.db.commit()

        if(signal):
            self.statusBar().showMessage('publisher removed', 3000)
            self.Show_Publisher()
        else:
            self.statusBar().showMessage('must be keyword mistake', 3000)


    def Delete_Category(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        delcat = self.lineEdit_24.text()

        sql = '''DELETE FROM category WHERE category_name=%s'''

        removed = self.cur.execute(sql,([delcat]))
        self.db.commit()
        if  removed :
            self.statusBar().showMessage('category removed', 3000)
            self.Show_Category()

        else:
            self.statusBar().showMessage('No category Found with that keyword')


    ######################################
    #######SETTINGS_SHOW_DATA_IN _COMBO###

    def Combo_show_Category(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT category_name FROM category''')

        data = self.cur.fetchall()

        self.comboBox_3.clear()
        self.comboBox_9.clear()

        for category in data:
            self.comboBox_3.addItem(category[0])
            self.comboBox_9.addItem(category[0])

    def Combo_show_Author(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT author_name FROM author''')

        data = self.cur.fetchall()

        self.comboBox_4.clear()
        self.comboBox_10.clear()

        for category in data:
            self.comboBox_4.addItem(category[0])
            self.comboBox_10.addItem(category[0])

    def Combo_show_Publisher(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT publisher_name FROM publisher''')

        data = self.cur.fetchall()
        self.comboBox_5.clear()
        self.comboBox_11.clear()

        for category in data:
            self.comboBox_5.addItem(category[0])
            self.comboBox_11.addItem(category[0])

        ######################################
        ##############Clients################

    def Add_Client(self):

        cname = self.lineEdit_5.text()

        cnationality = self.lineEdit_37.text()
        cemail = self.lineEdit_38.text()

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''
        INSERT INTO client(Client_name , Client_email , Client_nationality)
        VALUES (%s,%s,%s)''', (cname, cemail, cnationality))

        self.db.commit()
        self.db.close()

        self.statusBar().showMessage('                  New client added', 5000)

        self.lineEdit_5.setText('')
        self.lineEdit_37.setText('')
        self.lineEdit_38.setText('')

        self.Show_All_Client()

    def Search_client(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        searchC = self.lineEdit_28.text()

        sql = '''SELECT * FROM client WHERE Client_nationality = %s'''

        self.cur.execute(sql, ([searchC]))

        self.db.commit()

        data = self.cur.fetchone()

        if (data):

            self.lineEdit_25.setText(data[1])
            self.lineEdit_26.setText(data[2])
            self.lineEdit_27.setText(data[3])

        else:
            self.statusBar().showMessage('              No client Found !!', 4000)

    def Edit_Client(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        nameC = self.lineEdit_25.text()
        emailC = self.lineEdit_26.text()
        idC = self.lineEdit_27.text()
        idCS = self.lineEdit_28.text()

        self.cur.execute(
            '''UPDATE client SET Client_name=%s,Client_email=%s,Client_nationality=%s WHERE Client_nationality = %s'''
            , (nameC, emailC, idC, idCS))

        self.db.commit()
        self.statusBar().showMessage('              your data has been updated', 4000)

        self.Show_All_Client()

    def Delete_Client(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        deleteC = self.lineEdit_28.text()

        warning = QMessageBox.warning(self, 'Remove Client', 'Are you sure you want to remove client',
                                      QMessageBox.Yes | QMessageBox.No)

        if warning == QMessageBox.Yes:
            sql = '''DELETE FROM client where Client_nationality =%s'''

            self.cur.execute(sql, ([deleteC]))
            self.db.commit()
            self.db.close()

        self.statusBar().showMessage('              Client Removed', 4000)

        self.Show_All_Client()

    def Show_All_Client(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT Client_name , Client_email , Client_nationality FROM client ''')

        data = self.cur.fetchall()

        self.tableWidget_6.setRowCount(0)
        self.tableWidget_6.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget_6.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget_6.rowCount()
            self.tableWidget_6.insertRow(row_position)


    def Dayto_Day_Add(self):

        title_ = self.lineEdit.text()
        type_B = self.comboBox.currentText()
        days = self.comboBox_2.currentIndex() + 1
        client = self.lineEdit_6.text()
        date_t = datetime.date.today()
        to = date_t + datetime.timedelta(days=int(days))

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        sql = '''SELECT * FROM book WHERE book_name = %s'''
        book_avail = self.cur.execute(sql, ([title_]))

        sql2 ='''SELECT * FROM client WHERE Client_name =%s'''
        valid_client =self.cur.execute(sql2, ([client]))



        self.db.commit()


        if  book_avail:
            if valid_client:


                self.cur.execute('''INSERT INTO dayoperations (bookname,type_,days,client,DATETIME,to_)
                                    VALUES(%s,%s,%s,%s,%s,%s)''',
                                 (title_ ,type_B ,days ,client ,date_t ,to))

                self.db.commit()
                self.statusBar().showMessage('Done')
            else :
                self.statusBar().showMessage('No client with this name', 4000)

        else:
            self.statusBar().showMessage('Currently book not available', 4000)
        self.Show_All_Dattoday()

    def Show_All_Dattoday(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        title_ = self.lineEdit.text()

        self.cur.execute('''SELECT  bookname,type_,days,client,DATETIME,to_  FROM dayoperations''')



        data = self.cur.fetchall()

        self.tableWidget.setRowCount(0)
        self.tableWidget.insertRow(0)

        for row, form in enumerate(data):
            for column, item in enumerate(form):
                self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                column += 1

            row_position = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row_position)


        ########Export data to excel#####
        #################################

    def Export_Dailytdata(self):


        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT  bookname,type_,days,client,DATETIME,to_  FROM dayoperations''')

        data =self.cur.fetchall()

        wb = Workbook('DaytoDayoperations.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0,'bookname')
        sheet1.write(0,1,'type')
        sheet1.write(0,2,'days')
        sheet1.write(0,3,'client')
        sheet1.write(0,4,'To-Date')
        sheet1.write(0,5,'FROM-Date')


        row_number = 1
        for row in data:
            coulumn_number =0
            for item in row:
                sheet1.write(row_number,coulumn_number,str(item))
                coulumn_number +=1

            row_number+=1
        self.statusBar().showMessage('File created in your folder',3000)
        wb.close()
        self.statusBar().showMessage('Data Exported in  excel_sheet form', 3000)

    def Export_Bookdata(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()


        self.cur.execute('''SELECT book_name, book_description, book_code, book_category, book_author, book_publisher, book_price FROM book''')



        data =self.cur.fetchall()

        wb =Workbook("Books_details.xlsx")
        sheet2 = wb.add_worksheet()

        sheet2.write(0,0,'Book Name')
        sheet2.write(0, 1, 'Book Description')
        sheet2.write(0, 2, 'Book Code')
        sheet2.write(0, 3, 'Book Category')
        sheet2.write(0, 4, 'Book Auther')
        sheet2.write(0,5,'Book Publisher')
        sheet2.write(0,6,'Book Price')


        Row_number =1
        for row in data :
            Column_number = 0
            for item in row :

                sheet2.write(Row_number,Column_number,str(item))
                Column_number +=1

            Row_number +=1

        wb.close()
        self.statusBar().showMessage('Book Data saved in excel_sheet form', 3000)

    def Export_Clientdata(self):


        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()


        self.cur.execute('''SELECT Client_name, Client_email, Client_nationality  FROM  client''')


        data =self.cur.fetchall()

        wb = Workbook('Clients.xlsx')

        sheet3 = wb.add_worksheet()

        sheet3.write(0,0,'Client name')
        sheet3.write(0, 1, 'Client EMail')
        sheet3.write(0, 2, 'Client Nationality')


        Row_number =1
        for row in  data :

            Column_number =0
            for item in row :

                sheet3.write(Row_number,Column_number,str(item))
                Column_number+=1

            Row_number+=1


        wb.close()
        self.statusBar().showMessage('Client Data saved in excel_sheet form', 3000)

    def Export_User_data(self):

        self.db = MySQLdb.connect(host='localhost', user='root', password='', db='library')
        self.cur = self.db.cursor()




        self.cur.execute('''SELECT user_name,user_password,user_email  FROM  users''')

        data = self.cur.fetchall()

        wb = Workbook('Users_Data.xlsx')

        sheet3 = wb.add_worksheet()

        sheet3.write(0, 0, 'User Name')
        sheet3.write(0, 1, 'User Password')
        sheet3.write(0, 2, 'User Email')

        Row_number = 1
        for row in data:

            Column_number = 0
            for item in row:
                sheet3.write(Row_number, Column_number, str(item))
                Column_number += 1

            Row_number += 1

        wb.close()
        self.statusBar().showMessage('User Data saved in excel_sheet form',3000)


    def Admin_log(self):


        login = self.lineEdit_30.text()

        if login == '23654345676544567' :

            self.groupBox_5.setEnabled(True)




        ########Selecting Theme##########
        #################################
    def Theme_Orange(self):

        style =open('themes/darkorange.css' , 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Theme_Blue(self):

        style = open('themes/darkblue.css' , 'r')
        style = style.read()
        self.setStyleSheet(style)

    def Theme_Green(self):

        style = open('themes/darkbreeze.css','r')
        style = style.read()
        self.setStyleSheet(style)

    def Theme_Q(self):

        style = open('themes/Darkgrey.css' , 'r')
        style = style.read()
        self.setStyleSheet(style)



def main():
    app = QApplication(sys.argv)
    window = LoginUser()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()
