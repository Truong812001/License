#!/usr/bin/env python
# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QFileDialog, QMessageBox,QApplication,QInputDialog,QLineEdit
import openpyxl
import os
import shutil
import subprocess
import psse35
import psspy
from psspy import _i,_f,_s
import requests
from bs4 import BeautifulSoup
import re
from datetime import datetime
PSSE_PATH = r"C:\Program Files\PTI\PSSE35\35.4\PSSBIN"
import sys
import math
import pandas as pd 
import json
sys.path.append(PSSE_PATH)
os.environ['PATH'] += ';' + PSSE_PATH

class LicenseCheckDialog():
    def get_data_license(self):
        url = 'https://github.com/Truong812001/License/blob/main/License.txt'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        try:
            response = requests.get(url, headers=headers)    
            if response.status_code == 200:
                response = requests.get(url)
                response.raise_for_status()

                # Phân tích cú pháp HTML
                soup = BeautifulSoup(response.text, "html.parser")

                # Tìm thẻ <script> chứa thông tin JSON
                script_tag = soup.find("script", {"type": "application/json", "data-target": "react-app.embeddedData"})
                
                if script_tag:
                    # Tải nội dung JSON trong thẻ <script>
                    data = json.loads(script_tag.string)
                    
                    # Truy cập vào dữ liệu cần thiết
                    raw_lines = data.get("payload", {}).get("blob", {}).get("rawLines", [])
                    if raw_lines:
                        license_data = raw_lines[0]
                        update = raw_lines[1]
                        print("end date :", license_data)     
        except:
            return 9999999999,'no'
        return license_data, update
    def check_license(self):
        value = self.get_time_web()
        date_license,update = self.get_data_license()
        date_license = int(date_license)
        # Kiểm tra license
        if int(value) > date_license:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("License Error")
            msg_box.setText("Call Travis Right Now")
            msg_box.exec_()
            return False
        if update =='yes':
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setWindowTitle("UPDATE")
            msg_box.setText("New upload, Do you want to download?")
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            choice = msg_box.exec_()
            if choice == QMessageBox.Yes:
                folder = QFileDialog.getExistingDirectory(None, "Chọn thư mục để tải về")
                if folder:
                    print(f"Thư mục đã chọn: {folder}")
                    target_folder = os.path.join(folder, "Earn Money")
                    os.makedirs(target_folder, exist_ok=True)
                    print(f"Thư mục tải về: {target_folder}")
                    self.download_update(target_folder)
        return True
    def download_update(self,output_folder):
        api_url = "https://api.github.com/repos/Truong812001/License/contents/"
        try:
            # Gửi yêu cầu GET đến API
            response = requests.get(api_url)
            response.raise_for_status()
            files = response.json()

            # Tạo thư mục nếu chưa có
            os.makedirs(output_folder, exist_ok=True)

            # Tải từng file trong thư mục
            for file in files:
                if file['type'] == 'file':
                    file_url = file['download_url']
                    file_name = os.path.join(output_folder, file['name'])
                    
                    print(f"Tải file: {file_name}")
                    file_response = requests.get(file_url)
                    file_response.raise_for_status()
                    
                    # Lưu file vào thư mục
                    with open(file_name, 'wb') as f:
                        f.write(file_response.content)
        except requests.exceptions.RequestException as e:
            print(f"Lỗi khi tải thư mục: {e}")
        return 'Download Done!!'
    def prompt_password(self):
        password, ok = QInputDialog.getText(None, "Enter Password", "Nhập mật mã để vào ứng dụng:", QLineEdit.Password)
        if ok and password == "your_secret_password":  # Thay "your_secret_password" bằng mật mã bạn muốn
            return True
        else:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("Access Denied")
            msg_box.setText("Mật mã không chính xác. Không thể truy cập vào ứng dụng.")
            msg_box.exec_()
            return False
    def get_time_web(self):
        # Gửi yêu cầu GET đến trang web
        url = 'https://www.timeanddate.com/worldclock/vietnam/hanoi'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)

        # Kiểm tra xem yêu cầu có thành công không
        if response.status_code == 200:
            try:
                soup = BeautifulSoup(response.text, 'html.parser')
                date_string = soup.find('span', {'id': 'ctdat'}).text.strip()
                date_string = date_string.replace(' p.',"")
                if "Tháng" in date_string or "thứ" in date_string:
                    vietnamese_to_english = {
                        "thứ hai": "Monday", "thứ ba": "Tuesday", "thứ tư": "Wednesday",
                        "thứ năm": "Thursday", "thứ sáu": "Friday", "thứ bảy": "Saturday",
                        "chủ nhật":"Sunday",
                        "chúa nhật": "Sunday", "Tháng một": "January", "Tháng hai": "February",
                        "Tháng ba": "March", "Tháng tư": "April", "Tháng năm": "May",
                        "Tháng sáu": "June", "Tháng bảy": "July", "Tháng tám": "August",
                        "Tháng chín": "September",
                        "Tháng mười một": "November", "Tháng mười hai": "December","Tháng mười": "October"
                    }
                    for vi, en in vietnamese_to_english.items():
                        date_string = date_string.replace(vi, en)
                    date_string = date_string.strip()
                    date_format = "%A %d %B %Y"
                else:
                    # Nếu ngày tháng đã ở định dạng tiếng Anh
                    date_format = "%A %d %B %Y"

                # Chuyển đổi chuỗi sang datetime
                date = datetime.strptime(date_string, date_format)
                # Xuất ngày, tháng, và năm
                value = f"{date.year}{date.month:02}{date.day:02}"
            except:
                date_string = 10000000000
                print('something went wrong')
                value = date_string
        return value

class LeadLagWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(LeadLagWindow, self).__init__(parent)
        self.parent_window = parent
        self.setWindowTitle("Lead Lag Setup")
        self.parameters = {
            "Pnet": "",
            "Qnet":"0",
            "Delta P": "0.00001",
            "Delta Q": "0.03",
            "BranchPV": "",
            "BranchBESS": "",
            "PQCurve":"",
            "Scale Gen":""
        }
        self.resize(600, 500)  # Kích thước cửa sổ mới

        # Tính toán để đặt cửa sổ chính giữa
        if parent is not None:
            # Lấy tọa độ và kích thước của cửa sổ chính (parent)
            parent_geometry = parent.geometry()
            x = parent_geometry.x() + (parent_geometry.width() - self.width()) // 2
            y = parent_geometry.y() + (parent_geometry.height() - self.height()) // 2
            self.move(x, y)  # Đặt vị trí cửa sổ mới
        # Thêm ô nhập liệu (QLineEdit)
        self.input_field = QtWidgets.QLineEdit(self)
        self.input_field.setGeometry(190, 10, 350, 30)
        self.input_field.setPlaceholderText("Input path sav...")  # Gợi ý nhập liêuk

        self.resultstext = QtWidgets.QTextEdit(self)
        self.resultstext.setGeometry(QtCore.QRect(190, 50, 350, 400))
        self.resultstext.setObjectName("resultstext")

        # Thêm nút để xử lý dữ liệu từ ô nhập liệ
        self.function_button()
        self.function_action_button()
    def function_button(self):

        self.input_button = QtWidgets.QPushButton("Input", self)
        self.input_button.setGeometry(20, 15, 130, 31)

        self.set_up_param_button = QtWidgets.QPushButton("Set up Parameters", self)
        self.set_up_param_button.setGeometry(20, 60, 130, 31)

        self.basic_case_button = QtWidgets.QPushButton("BASIC CASE", self)
        self.basic_case_button.setGeometry(20, 105, 130, 31)

        self.toolP_button = QtWidgets.QPushButton("Tool P", self)
        self.toolP_button.setGeometry(20, 150, 130, 31)

        self.toolV_button = QtWidgets.QPushButton("Tool V", self)
        self.toolV_button.setGeometry(20, 195, 130, 31)

        self.lead_lag_button = QtWidgets.QPushButton("LEAD LAG", self)
        self.lead_lag_button.setGeometry(20, 240, 130, 31)

        self.lead_lag_auto_button = QtWidgets.QPushButton("Auto Tap", self)
        self.lead_lag_auto_button.setGeometry(20, 295, 130, 31)
    def function_action_button(self):
        self.input_button.clicked.connect(lambda: self.openFileDialog(self.input_button))
        self.set_up_param_button.clicked.connect(self.set_up_param)
        self.basic_case_button.clicked.connect(self.basic_case)
        self.toolP_button.clicked.connect(self.ToolP)
        self.toolV_button.clicked.connect(self.ToolV)
        self.lead_lag_button.clicked.connect(self.Lead_lag)
        self.lead_lag_auto_button.clicked.connect(self.auto_Lead_lag)

    def ToolP(self):
        path = self.input_field.text()
        if not path:
            self.msg_box("file")
        else:
            try:
                param = self.param_set_up
                tool_lead_lag = TOOL_LEAD_LAG(param,path)
                ierr = psspy.psseinit()
                ierr = psspy.case(path)
                res =tool_lead_lag._adjust_active_power(float(param['Pnet']))
                if type(res) is str:
                    self.resultstext.append(res)
                else:
                    self.resultstext.append(f"P each gen:{res}")
            except Exception as e:
                self.resultstext.append(str(e))

    def ToolV(self):
        path = self.input_field.text()
        if not path:
            self.msg_box("file")
        else:
            try:
                param = self.param_set_up
                tool_lead_lag = TOOL_LEAD_LAG(param,path)
                ierr = psspy.psseinit()
                ierr = psspy.case(path)
                res =tool_lead_lag._adjust_reactive_power(float(param['Qnet']))
                if type(res) is str:
                    self.resultstext.append(res)
                else:
                    self.resultstext.append(f"Vsch:{res}")
            except Exception as e:
                self.resultstext.append(str(e))
    def basic_case(self):
        path = self.input_field.text()
        if not path:
            self.msg_box("file")
        else:
            try:
                param = self.param_set_up
                ## path is Sav file
                tool_lead_lag = TOOL_LEAD_LAG(param,path)
                if param.get('BranchPV') !="" and  param.get('BranchBESS') !="":
                    self.resultstext.append("Tune PV + BESS")
                    res = tool_lead_lag.PV_BESS(float(param.get('Pnet')),float(param.get('Qnet')))
                    print(param['BranchBESS'])
                    if type(res) is str :
                        self.resultstext.append(res)
                    else:
                        self.resultstext.append(f"PV\n Pgen:{res[0][0]}\n  Vsch:{res[0][1]}   \n BESS\n Pgen:{res[1][0]}\n  Vsch:{res[1][1]}")
                elif param.get('BranchPV') == "1":
                    self.resultstext.append("Tune PV Alone")
                    res = tool_lead_lag.PV_alone(float(param.get('Pnet')),float(param.get('Qnet')))
                    print('pv alone')
                    if type(res) is str :
                        self.resultstext.append(res)
                    else:
                        self.resultstext.append(f"Pgen\n{res[0]}\nVsch\n{res[1]}\nDONE!!!")
                elif param.get("BranchBESS") =="1":
                    print("bess alone")
                    self.resultstext.append("Tune BESS Alone")
                    res  = tool_lead_lag.BESS_alone(float(param.get('Pnet')),float(param.get('Qnet')))
                    if type(res) is str :
                        self.resultstext.append(res)
                    else:
                        self.resultstext.append(f"BESSD\nPgen\n{res[0]}\nVsch\n{res[1]}\nDONE!!!\n\nBESSC\nPgen\n{res[3]}\nVsch\n{res[4]}\nDONE\n!!!")

                elif param.get('BranchPV') =="" and  param.get('BranchBESS') =="":
                    self.resultstext.append('No data branch Fail!!!!')
                    print('ok')
            except Exception as e:
                print(str(e))
                return self.resultstext.append("Set Up and Basic Case Fail!!!")
    def Lead_lag(self):
        path = self.input_field.text()
        if not path:
            self.msg_box("file")
        else:
            try:
                param = self.param_set_up
                ## path is Sav file
                tool_lead_lag = TOOL_LEAD_LAG(param,path)                       
                Vsch_lag,Vsch_lead = tool_lead_lag.LEAD_LAG(float(param['Pnet']))
                self.resultstext.append(f'Vsch LAG : {Vsch_lag}\nVsch LEAD : {Vsch_lead}')
            except Exception as e:
                print(str(e))
                return self.resultstext.append(str(e))
    def auto_Lead_lag(self):
        path = self.input_field.text()
        if not path:
            self.msg_box("file")
        else:
            reply = self.messagebox('Run')
            if reply == QMessageBox.Ok:
                try:
                    param = self.param_set_up
                    ## path is Sav file
                    tool_lead_lag = TOOL_LEAD_LAG(param,path)                       
                    res = tool_lead_lag.auto_Lead_Lag(float(param['Pnet']))
                    # res = {'auto_LAG': {1.045: 1.05893, 1.06: 1.04756, 1.075: 1.03659, 1.09: 1.025991}, 'auto_LEAD':{}}

                    if type(res) is str:
                        return self.resultstext.append(res)
                    self.resultstext.append("DONE!!! Please check excel file")

                    if res['auto_LEAD']=={} :
                       lead_results = ['NO RESULTS'] 
                    else:
                        lead_ = res.get('auto_LEAD').keys()
                        lead_results = list(map(str, lead_)) 

                    if res['auto_LAG']=={} :
                       lag_results = ['NO RESULTS'] 
                    else:
                        lag_ = res.get('auto_LAG').keys()
                        lag_results = list(map(str, lag_))                        

                    lag_ratio,lead_ratio  = self.show_result_selection(lag_results,lead_results)
                    self.resultstext.append(f'LAG ratio: {lag_ratio}\nLAG Vsch: {res["auto_LAG"][float(lag_ratio)]}\nLEAD ratio: {lead_ratio}\nLEAD Vsch:{res["auto_LEAD"][float(lead_ratio)]}')
                    if lag_ratio == None or lead_ratio == None:
                        return 
                    
                    new_folder = os.path.join(tool_lead_lag.folder, '_auto_LEADLAG')
                    if os.path.exists(new_folder):
                        shutil.rmtree(new_folder)
                    os.makedirs(new_folder, exist_ok=True)
                    if lag_ratio != 'NO RESULTS':
                        hv_ = os.path.join(new_folder, os.path.basename(tool_lead_lag.file))
                        hv_path = os.path.join(new_folder, "project_lag_hv.sav" )
                        shutil.copy(path, new_folder) 
                        os.rename(hv_,hv_path)
                        ierr = psspy.psseinit()
                        ierr = psspy.case(hv_path) 
                        ierr, bus = psspy.atr3int(sid = -1 , flag=1, entry = 1, string = ['WIND1NUMBER','WIND2NUMBER','WIND3NUMBER'] )
                        ierr, ID  = psspy.atr3char(sid = -1 , flag=1, entry = 1, string = ['ID'])
                        gen_number,poi_number,id_gen,id_poi = tool_lead_lag.count_gen()       
                        tool_lead_lag._update_voltage_schedule(gen_number,res["auto_LAG"][float(lag_ratio)])
                        tool_lead_lag._on_auto_tap_update_ratio(bus,ID,float(lag_ratio))
                        tool_lead_lag._run_power_flow()
                        psspy.save(hv_path)

                        lv_path = os.path.join(new_folder, "project_lag_lv.sav")
                        shutil.copy(hv_path, lv_path)                            
            
                    if lead_ratio != 'NO RESULTS':
                        hv1_ = os.path.join(new_folder, os.path.basename(tool_lead_lag.file))
                        hv1_path = os.path.join(new_folder, "project_lead_hv.sav" )
                        shutil.copy(path, new_folder) 
                        os.rename(hv1_,hv1_path)
                        ierr = psspy.psseinit()
                        ierr = psspy.case(hv1_path) 
                        warn = tool_lead_lag.off_shunt()
                        ierr, bus = psspy.atr3int(sid = -1 , flag=1, entry = 1, string = ['WIND1NUMBER','WIND2NUMBER','WIND3NUMBER'] )
                        ierr, ID  = psspy.atr3char(sid = -1 , flag=1, entry = 1, string = ['ID'])
                        gen_number,poi_number,id_gen,id_poi = tool_lead_lag.count_gen()       
                        tool_lead_lag._update_voltage_schedule(gen_number,res["auto_LEAD"][float(lead_ratio)])
                        tool_lead_lag._on_auto_tap_update_ratio(bus,ID,float(lead_ratio))
                        tool_lead_lag._run_power_flow()
                        psspy.save(hv1_path)

                        lv1_path = os.path.join(new_folder, "project_lead_lv.sav")
                        shutil.copy(hv1_path, lv1_path)                       
                    

                except Exception as e:
                    print(str(e))
                    return self.resultstext.append(str(e))  
    def show_result_selection(self, lag_results,lead_results):
        # Tạo cửa sổ mới
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("TAP")
        dialog.resize(300, 200)

        # Thêm nhãn hướng dẫn
        lag = QtWidgets.QLabel("LAG:", dialog)
        lag.setGeometry(10, 10, 280, 30)

        lead = QtWidgets.QLabel("LEAD:", dialog)
        lead.setGeometry(10, 80, 280, 30)
        # Tạo QComboBox để hiển thị danh sách kết quả
        lag_box = QtWidgets.QComboBox(dialog)
        lag_box.setGeometry(10, 50, 280, 30)
        lag_box.addItems(lag_results)  # Thêm danh sách kết quả vào combo box

        lead_box = QtWidgets.QComboBox(dialog)
        lead_box.setGeometry(10, 120, 280, 30)
        lead_box.addItems(lead_results)
        # Nút OK để xác nhận
        ok_button = QtWidgets.QPushButton("OK", dialog)
        ok_button.setGeometry(100, 160, 100, 30)
        ok_button.clicked.connect(dialog.accept)

        # Hiển thị cửa sổ và lấy giá trị được chọn khi nhấn OK
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            lag = lag_box.currentText()
            lead = lead_box.currentText()
            # self.resultstext.append(f"LAG Ratio: {lag}\nLEAD Ratio: {lead}")   
            return lag,lead   
        else:
            return None ,None
    def set_up_param(self):
        file = self.input_field.text()
        if not file:  # Nếu nội dung rỗng
            msg_box = QtWidgets.QMessageBox(self)
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
            msg_box.setWindowTitle("Error")
            msg_box.setText("No file detected. Please input a valid path.")
            msg_box.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg_box.exec_()
        else:
            self.param_set_up = self.open_input_dialog()
            return self.param_set_up
    def msg_box(self,file):
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Error")
            msg_box.setText(f"No {file} selected.")
            msg_box.exec_()
    def open_input_dialog(self):
        # Mở hộp thoại nhập thông số
        dialog = InputDialog(self,self.parameters)
        if dialog.exec_():  # Chỉ thực hiện khi nhấn OK
            self.parameters = dialog.get_parameters()
            self.resultstext.clear()
            for key, value in self.parameters.items():
                self.resultstext.append(f"{key}: {value}")
            return self.parameters
    def openFileDialog(self, button):
        # Mở hộp thoại chọn thư mục
        options = QtWidgets.QFileDialog.Options()
        if button == self.input_button:
            file_filter = "SAV Files (*.sav);;All Files (*)"
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select File", "", file_filter, options=options)

        # Đặt đường dẫn vào QLineEdit tương ứng
        if button == self.input_button:
            self.input_field.setText(fileName)

    def closeEvent(self, event):
        # Khi cửa sổ LeadLagWindow bị đóng, hiển thị lại cửa sổ chính
        if self.parent_window:
            self.parent_window.show()
        super().closeEvent(event)
    def messagebox(self,Name):
        reply = QMessageBox.question(
        None, Name, 'Start??',
        QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Cancel
            )
        return reply

class InputDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, default_values=None):
        super(InputDialog, self).__init__(parent)
        self.setWindowTitle("Enter Parameters")
        self.resize(500, 300)
        self.default_values = default_values or {}
        # Layout chính
        layout = QtWidgets.QVBoxLayout(self)

        # Nhập tham số 1
        self.Pnet_label = QtWidgets.QLabel("Pnet:")
        self.Pnet_input = QtWidgets.QLineEdit(self)
        self.Pnet_input.setPlaceholderText("Example: 100 or -100")
        self.Pnet_input.setText(self.default_values.get("Pnet", ""))
        layout.addWidget(self.Pnet_label)
        layout.addWidget(self.Pnet_input)

        self.Qnet_label = QtWidgets.QLabel("Qnet:")
        self.Qnet_input = QtWidgets.QLineEdit(self)
        self.Qnet_input.setText(self.default_values.get("Qnet", ""))
        layout.addWidget(self.Qnet_label)
        layout.addWidget(self.Qnet_input)

        # Nhập tham số 2
        self.DeltaP_label = QtWidgets.QLabel("Delta P:")
        self.DeltaP_input = QtWidgets.QLineEdit(self)

        self.DeltaP_input.setText(self.default_values.get("Delta P", "0.00001"))
        layout.addWidget(self.DeltaP_label)
        layout.addWidget(self.DeltaP_input)


        self.DeltaQ_label = QtWidgets.QLabel("Delta Q:")
        self.DeltaQ_input = QtWidgets.QLineEdit(self)

        self.DeltaQ_input.setText(self.default_values.get("Delta Q", "0.03"))
        layout.addWidget(self.DeltaQ_label)
        layout.addWidget(self.DeltaQ_input)


        self.BranchPV_label = QtWidgets.QLabel("Branch of PV:")
        self.BranchPV_input = QtWidgets.QLineEdit(self)
        self.BranchPV_input.setPlaceholderText("Example: 10001,11000,12000,... Only: 1")
        self.BranchPV_input.setText(self.default_values.get("BranchPV", ""))
        layout.addWidget(self.BranchPV_label)
        layout.addWidget(self.BranchPV_input)

        self.BranchBESS_label = QtWidgets.QLabel("Branch of BESS:")
        self.BranchBESS_input = QtWidgets.QLineEdit(self)
        self.BranchBESS_input.setPlaceholderText("Example: 10001,11000,12000,... Only: 1")
        self.BranchBESS_input.setText(self.default_values.get("BranchBESS", ""))
        layout.addWidget(self.BranchBESS_label)
        layout.addWidget(self.BranchBESS_input)

        self.OptionPQcurve_label = QtWidgets.QLabel("PQCurve:")
        self.OptionPQcurve_input = QtWidgets.QComboBox(self)
        self.OptionPQcurve_input.addItems(["Yes", "No"])
        # Đặt giá trị mặc định
        default_pqcurve = "No" if self.default_values.get("PQCurve", "").lower() == "yes" else "No"
        self.OptionPQcurve_input.setCurrentText(default_pqcurve)
        layout.addWidget(self.OptionPQcurve_label)
        layout.addWidget(self.OptionPQcurve_input)

        self.OptionScaleGen_label = QtWidgets.QLabel("Scale Gen Follow MVABase:")
        self.OptionScaleGen_input = QtWidgets.QComboBox(self)
        self.OptionScaleGen_input.addItems(["Yes", "No"])
        # Đặt giá trị mặc định
        default_ScaleGen = "No" if self.default_values.get("ScaleGen", "").lower() == "yes" else "No"
        self.OptionScaleGen_input.setCurrentText(default_ScaleGen)
        layout.addWidget(self.OptionScaleGen_label)
        layout.addWidget(self.OptionScaleGen_input)
        # Nút OK và Cancel
        self.button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def get_parameters(self):
        # Lấy giá trị nhập từ các ô

        Pnet = self.Pnet_input.text()
        Qnet = self.Qnet_input.text()
        DeltaP = self.DeltaP_input.text()
        DeltaQ = self.DeltaQ_input.text()
        BranchPV= self.BranchPV_input.text()
        BranchBESS= self.BranchBESS_input.text()
        PQCurve = self.OptionPQcurve_input.currentText()
        ScaleGen = self.OptionScaleGen_input.currentText()
        return {"Pnet": Pnet,"Qnet": Qnet, "Delta P": DeltaP, "Delta Q": DeltaQ,"BranchPV": BranchPV,"BranchBESS": BranchBESS,"PQCurve": PQCurve,"ScaleGen": ScaleGen}

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        # Đặt tên đối tượng cho cửa sổ chính
        MainWindow.setObjectName("PYPSA")

        # Thiết lập kích thước của cửa sổ chính
        MainWindow.resize(600,500)


        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("bk1.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        ## ico

        MainWindow.setWindowIcon(QtGui.QIcon('bk1.png'))
        # Tạo một widget trung tâm cho cửa sổ chính
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        ##input
        ## creat button tạo nút bấm
        self.function_create_button()

        # tạm thời Tạo các hành động cho menu (chưa tạo menu)
        self.function_creat_sub_menu()

        self.resultsTextEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.resultsTextEdit.setGeometry(QtCore.QRect(190, 50, 350, 400))
        self.resultsTextEdit.setObjectName("resultsTextEdit")
        # Đặt widget trung tâm cho cửa sổ chính
        MainWindow.setCentralWidget(self.centralwidget)

        # Tạo một menu bar cho cửa sổ chính
        self.function_create_menu_bar()
        # Thêm các hành động vào menu Home
        self.function_create_sub_action_menu()

        self.retranslateUi(MainWindow)

        # Kết nối các tín hiệu và slot
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Kết nối các nút bấm với hàm mở hộp thoại chọn tệp
        ## run when click button
        self.function_button_action()

    def function_create_button(self):

        self.lineINPUT = QtWidgets.QLineEdit(self.centralwidget)
        self.lineINPUT.setGeometry(QtCore.QRect(190, 10, 350, 30))
        self.lineINPUT.setObjectName("lineINPUT")

        # Tạo các nút để mở hộp thoại chọn tệp tin
        self.buttonINPUT = QtWidgets.QPushButton(self.centralwidget)
        self.buttonINPUT.setGeometry(QtCore.QRect(30, 10, 100, 30))
        self.buttonINPUT.setObjectName("buttonINPUT")

        self.buttonDELETE_CNV = QtWidgets.QPushButton(self.centralwidget)
        self.buttonDELETE_CNV.setGeometry(QtCore.QRect(20, 60, 130, 31))
        self.buttonDELETE_CNV.setObjectName("buttonDELETE_CNV")

        self.buttonTOOLC = QtWidgets.QPushButton(self.centralwidget)
        self.buttonTOOLC.setGeometry(QtCore.QRect(20, 105, 130, 31))
        self.buttonTOOLC.setObjectName("buttonTOOLC")

        self.buttonTOOLD = QtWidgets.QPushButton(self.centralwidget)
        self.buttonTOOLD.setGeometry(QtCore.QRect(20, 150, 130, 31))
        self.buttonTOOLD.setObjectName("buttonTOOLD")

        self.buttonON_TAP = QtWidgets.QPushButton(self.centralwidget)
        self.buttonON_TAP.setGeometry(QtCore.QRect(20, 195, 130, 31))
        self.buttonON_TAP.setObjectName("buttonON_TAP")

        self.buttonSAVE_RAW = QtWidgets.QPushButton(self.centralwidget)
        self.buttonSAVE_RAW.setGeometry(QtCore.QRect(20, 250, 130, 31))
        self.buttonSAVE_RAW.setObjectName("buttonSAVE_RAW")

        self.buttonCHECK_LOG = QtWidgets.QPushButton(self.centralwidget)
        self.buttonCHECK_LOG.setGeometry(QtCore.QRect(20, 305, 130, 31))
        self.buttonCHECK_LOG.setObjectName("buttonCHECK_LOG")

    def function_create_menu_bar(self):
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 20, 486, 21))
        self.menubar.setObjectName("menubar")
        menu_bar_style = """
            QMenuBar {
                background-color: #D3D3D3;
                border: 1px solid black;
                color: black;
            }
        """
        self.menubar.setStyleSheet(menu_bar_style)
        # Tạo các mục menu
        self.menuHome = QtWidgets.QMenu(self.menubar)
        self.menuHome.setObjectName("menuHome")
##        self.menuHome.setObjectName("menuHelp")

        self.menuSimulate = QtWidgets.QMenu(self.menubar)
        self.menuSimulate.setObjectName("menuSimulate")

        self.menucheck_list = QtWidgets.QMenu(self.menubar)
        self.menucheck_list.setObjectName("menucheck_list")
        # Đặt menu bar cho cửa sổ chính
        MainWindow.setMenuBar(self.menubar)

        # Tạo một status bar cho cửa sổ chính
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        ##add menu home cho taskbar
        self.menubar.addAction(self.menuHome.menuAction())
        self.menubar.addAction(self.menuSimulate.menuAction())
        self.menubar.addAction(self.menucheck_list.menuAction())
    def function_create_sub_action_menu(self):
        self.menuHome.addAction(self.actionReportExcelMQT)
        self.menuSimulate.addAction(self.actionSetupLeadLag)
        self.menucheck_list.addAction(self.actionCheckBasicModel)
        self.menucheck_list.addAction(self.actionCheckLeadLagModel)
        self.menucheck_list.addAction(self.actionCheckALL)

    def function_creat_sub_menu(self):
        self.actionReportExcelMQT = QtWidgets.QAction(MainWindow)
        self.actionReportExcelMQT.setObjectName("actionReportExcelMQT")

        self.actionSetupLeadLag = QtWidgets.QAction(MainWindow)
        self.actionSetupLeadLag.setObjectName("actionSetupLeadLag")

        self.actionCheckBasicModel = QtWidgets.QAction(MainWindow)
        self.actionCheckBasicModel.setObjectName("actionCheckBasicModel")      

        self.actionCheckLeadLagModel = QtWidgets.QAction(MainWindow)
        self.actionCheckLeadLagModel.setObjectName("actionCheckLeadLagModel") 

        self.actionCheckALL = QtWidgets.QAction(MainWindow)
        self.actionCheckALL.setObjectName("actionCheckALL")
    def function_button_action(self):
        self.DMW = DMW(self)
        self.PSSE = PSSE(self)
        ## input
        self.buttonINPUT.clicked.connect(lambda: self.openFolderDialog(self.buttonINPUT))
        ##DMW
        self.buttonDELETE_CNV.clicked.connect(self.DMW.delete_cnv)
        self.buttonTOOLC.clicked.connect(lambda: self.DMW.run_dmw(self.DMW.dmwc()))
        self.buttonTOOLD.clicked.connect(lambda: self.DMW.run_dmw(self.DMW.dmwd()))
        self.buttonCHECK_LOG.clicked.connect(self.DMW.check_log)
        ##PSSE
        self.buttonSAVE_RAW.clicked.connect(self.PSSE.save_raw)
        self.buttonON_TAP.clicked.connect(self.PSSE.on_tap)

        ### connect button Report
        self.Report_MQT = Report_MQT(self)
        self.actionReportExcelMQT.triggered.connect(self.Report_MQT.main)

        ##lead lag
        self.actionSetupLeadLag.triggered.connect(self.openLeadLagWindow)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate

        MainWindow.setWindowTitle(_translate("MainWindow", "MQT TEAM"))
        MainWindow.setWindowIcon(QtGui.QIcon('bk1.png'))
        self.buttonINPUT.setText(_translate("MainWindow", "INPUT DMW"))
        self.buttonDELETE_CNV.setText(_translate("MainWindow", "DELETE_CNV"))
        self.buttonTOOLC.setText(_translate("MainWindow", "TOOLC"))
        self.buttonTOOLD.setText(_translate("MainWindow", "TOOLD"))
        self.buttonON_TAP.setText(_translate("MainWindow", "ON_TAP"))
        self.buttonSAVE_RAW.setText(_translate("MainWindow", "SAVE_RAW"))
        self.buttonCHECK_LOG.setText(_translate("MainWindow", "CHECK_LOG"))
        self.menuHome.setTitle(_translate("MainWindow", "Home"))
        self.actionReportExcelMQT.setText(_translate("MainWindow", "ReportExcelMQT"))


        self.menuSimulate.setTitle(_translate("MainWindow", "Simulate"))
        self.actionSetupLeadLag.setText(_translate("MainWindow", "Set up Lead Lag"))

        self.menucheck_list.setTitle(_translate("MainWindow", "Check list"))
        self.actionCheckBasicModel.setText(_translate("MainWindow", "Basic Model"))
        self.actionCheckLeadLagModel.setText(_translate("MainWindow", "Lead Lag Model"))
        self.actionCheckALL.setText(_translate("MainWindow", "Check All in DMV"))
##        self.menuOption.setTitle(_translate("MainWindow", "Option"))
        # self.menuHelp.setTitle(_translate("MainWindow", "Help"))

    def openFolderDialog(self, button):
        # Mở hộp thoại chọn thư mục
        options = QtWidgets.QFileDialog.Options()
        if button == self.buttonINPUT:
            folder_filter = "All Files (*)"
        folderName = QtWidgets.QFileDialog.getExistingDirectory(None, "Select Folder", "", options=options)

        # Đặt đường dẫn vào QLineEdit tương ứng
        if button == self.buttonINPUT:
            self.lineINPUT.setText(folderName)

    def openLeadLagWindow(self):
        MainWindow.hide()
        self.lead_lag_window = LeadLagWindow(MainWindow)

        self.lead_lag_window.show()

    def msg_box(self,file):
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Error")
            msg_box.setText(f"No {file} selected.")
            msg_box.exec_()

    def MessageBox(self,Name):
        reply = QMessageBox.question(
        None, Name, 'Start??',
        QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Cancel
            )
        return reply
def run_batch(bat_file,project):
    try:
        result = subprocess.run([bat_file, project], capture_output=True, text=True,creationflags=subprocess.CREATE_NO_WINDOW)
        return project, result.stdout, result.stderr, result.returncode
    except FileNotFoundError:
        return project, None, f"Không tìm thấy file batch: {bat_file}", -1
    except Exception as e:
        return project, None, str(e), -1

class DMW():
    def __init__(self,ui):
        self.ui = ui
    def delete_cnv(self):
        folder = self.ui.lineINPUT.text()
        self.ui.resultsTextEdit.clear()
        if  not folder:
            self.ui.msg_box("folder")
        else:
            path = os.path.join(folder, "CASEs", "project")
            if os.path.exists(path):  # Kiểm tra xem thư mục có tồn tại không
                for file in os.listdir(path):
                    if file.endswith("cnv.sav") or file.endswith(".snp") or file.endswith(".flx") or file.endswith(".log") or file.endswith("cnv.raw"):
                        try:
                            os.remove(os.path.join(path, file))
                            self.ui.resultsTextEdit.append(f"Deleted {file}")
                        except Exception as e:
                            # In thông báo lỗi vào QTextEdit nếu có lỗi khi xóa file
                            self.ui.resultsTextEdit.append(f"Error deleting {file}: {e}")
            else:
                self.ui.resultsTextEdit.append("Error: CASEs/project folder not found!")
            self.ui.resultsTextEdit.append('DELETE CNV DONE')
        return folder
    def check_log(self):
        self.ui.resultsTextEdit.clear()
        folder = self.ui.lineINPUT.text()
        if  not folder:
            self.ui.msg_box("folder")
        else:
            path =os.path.join(folder, "RESULTs")
            first_folder = True
            for root, dirs, files in os.walk(path):
        ##        print("Đường dẫn:", root)
                if first_folder :
                    first_folder = False
                    continue
                for file in files:
                    if file.endswith('.log'):
                        path_log = os.path.join(root, file)
                        with open(path_log, 'r') as f:
                            content = f.read()
                            if "O.K." in content:
                                self.ui.resultsTextEdit.append('INITIAL')
                            else:
                                self.ui.resultsTextEdit.append(f'NOT INITIAL: {file}!!!!!!!!!!!')
            self.ui.resultsTextEdit.append('CHECK LOG DONE')
    def dmwc(self):
        return ["cproject_lag_hv", "cproject_lag_lv", "cproject_lead_lv", "cproject_lead_hv","1cproject_lag_hv_prefer_pvbess2","1cproject_lag_lv_prefer_pvbess2","1cproject_lead_hv_prefer_pvbess2","1cproject_lead_lv_prefer_pvbess2"]
    def dmwd(self):
        return ["project_lag_hv", "project_lag_lv", "project_lead_lv", "project_lead_hv","2project_lag_hv_prefer_pvbess2","2project_lag_lv_prefer_pvbess2","2project_lead_hv_prefer_pvbess2","2project_lead_lv_prefer_pvbess2"]
    def run_dmw(self,args):
        self.ui.resultsTextEdit.clear()
        folder = self.ui.lineINPUT.text()
        if  not folder:
            self.ui.msg_box("folder")
##            reply = QMessageBox.question(self,'Start?','Do you want to start?',QMessageBox.Ok | QMessageBox.Cancel,QMessageBox.Cancel)
        else:
            reply = self.ui.MessageBox('DMView')

##            check = self.msg_box1("Start?")

            if reply == QMessageBox.Ok:

                os.chdir(folder)
                bat_file = r"run3x.bat"
                import time
                start_overall_time = time.time()
                from concurrent.futures import ThreadPoolExecutor, as_completed

                with ThreadPoolExecutor() as executor:
                    futures = {executor.submit(run_batch,bat_file,project): project for project in args}

                    for future in as_completed(futures):
                        project = futures[future]
                        try:
                            project, stdout, stderr, returncode = future.result()

                            project_name = os.path.basename(project)
                            result_text = f"Project: {project_name}\n"
                            print("STDOUT:")
                            print(stdout)
                            print("STDERR:")
                            print(stderr)

                            if returncode != 0:
                                result_text += stderr
                            else:
                                result_text += "DONE!\n"
                            self.ui.resultsTextEdit.append(result_text)
                            QApplication.processEvents()
                        except Exception as e:
                            print(f"Lỗi khi chạy project {project}: {str(e)}")
            else:
                return
            end_overall_time = time.time()
            overall_execution_time = end_overall_time - start_overall_time
##            self.ui.msg_box1(f"DONE {overall_execution_time}s")

class PSSE():
    def __init__(self,ui):
        self.ui = ui
    def save_raw(self):
        self.ui.resultsTextEdit.clear()
        folder = DMW(ui).delete_cnv()
        if  not folder:
            self.ui.msg_box("folder")
        else:
            path = os.path.join(folder, "CASEs", "project")
            if os.path.exists(path):
                for file_name in os.listdir(path):
                    print(file_name)
                    if file_name.endswith('.sav'):
                        full_path = os.path.join(path, file_name)
                        raw = file_name.replace(".sav", ".raw")
                        path_raw =os.path.join(path, raw)
                        ierr = psspy.psseinit()
                        ierr = psspy.case(full_path)
                        psspy.fnsl([1,1,0,0,1,1,0,0])
                        psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,path_raw)
                        psspy.save(full_path)
                        self.ui.resultsTextEdit.append(raw)
                        self.ui.resultsTextEdit.append('SAVE RAW FILE DONE')
                        with open(path_raw,"r") as file:
                            lines=file.readlines()
                        lines[1] = lines[1].replace(" Xplore", "")
                        with open(path_raw, 'w') as file:
                            file.writelines(lines)
                            self.ui.resultsTextEdit.append("DELETE XPLORE")
            else:
                self.ui.resultsTextEdit.append("Error",path)

    def on_tap(self):
        reply = self.ui.MessageBox('ONTAP')
        if reply == QMessageBox.Ok:
            self.ui.resultsTextEdit.clear()
            folder = DMW(ui).delete_cnv()
            if  not folder:
                self.ui.msg_box("folder")
            else:
                path = os.path.join(folder, "CASEs", "project")
                for file_name in os.listdir(path):
                    if file_name.endswith('.sav'):
                        full_path = os.path.join(path, file_name)
                        ierr = psspy.psseinit()
                        ierr = psspy.case(full_path)
                        ierr, bus = psspy.atr3int(sid = -1 , owner = 1, ties = 3, flag=2, entry = 1, string = ['WIND1NUMBER','WIND2NUMBER','WIND3NUMBER'] )
    ##                    ierr, ratio = pssy.awndreal(sid = -1, owner = 1 , ties, flag, entry, string)
                        ierr, ID  = psspy.atr3char(sid = -1 , flag=1, entry = 1, string = ['ID'])
                        print(bus)
                        if ierr ==0:
                            for i in range(len(bus[0])):
                                psspy.three_wnd_winding_data_5(bus[0][i],bus[1][i],bus[2][i],ID[0][i],1,[_i,_i,_i,_i,_i,1],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
                                psspy.fnsl([1,1,0,0,1,1,0,0])
    ##                            psspy.rawd_2(0,1,[1,1,1,0,0,0,0],0,path_raw)
                                psspy.save(full_path)
                        else:
                            self.ui.resultsTextEdit.append("SOMETHING ON TAP WENT WRONG!!")

                self.ui.resultsTextEdit.append("ON TAP DONE!!")

    def toolPV(self):
        return
    def toolBESS(self):
        return
class TOOL_LEAD_LAG:
    def __init__(self,param,path):
        ## độ dài số thập phân sau dấu phẩy
        self.MAX_ITER = 7
        ## sai số
        self.ERROR_TOLERANCE = float(param['Delta P'])
        self.ERROR_TOLERANCE1 = float(param['Delta Q'])
        self.PQCurve = param['PQCurve']
        self.file = path
        self.branchbess = param.get('BranchBESS')
        self.branchpv = param.get('BranchPV')
        self.ScaleGen = param['ScaleGen']
        ## sav in folder
        self.folder = os.path.dirname(path)
    def PV_BESS(self,Pnet,Qnet):
        if Pnet >= 0:
            return "TUNE PV PNET < 0!!!!!!!!"
        # if self.Set_up_for_BESS() or self.Set_up_for_PV():
        #     return " Branch BESS or PV ???"
        res1 = self.PV_alone(Pnet,Qnet)
        if type(res1) is str:
            res1 = res1 +' PV_alone'
            return res1
        else:
            P_each_gen_pv = res1[0]
            Vsch_pv = res1[1]
            path_pv = res1[2]

        res2 = self.BESS_alone(Pnet,Qnet)
        if type(res2) is str :
            res2 = res2 +' BESS alone'
            return res2
        else:
            P_each_gen_bessd = res2[0]
            Vsch_bessd = res2[1]
            path_bessd = res2[2]
            P_each_gen_bessc = res2[3]
            Vsch_bessc = res2[4]
            path_bessc = res2[5]
        ##update
        ## PValone set up  pv
        ## on bess off pv
        try:
            ##PV
            ierr = psspy.psseinit()
            ierr = psspy.case(path_pv)
            ## on bess to update bess
            self.Set_up_for_BESS()
            Pmax = P_each_gen_bessd
            Pmin = P_each_gen_bessc
            self._update_Qmax_Qmin_Pmax_Pmin_Pgen(P_each_gen_bessd,Pmax,Vsch_bessd,Pmin)
            # off pv to update value PV
            self.Set_up_for_PV()
            self._run_power_flow()
            psspy.save(path_pv)


            ## BESSD
            ierr = psspy.psseinit()
            ierr = psspy.case(path_bessd)
            self.Set_up_for_PV()
            Pmin = [0] * len(P_each_gen_pv)
            Pmax = P_each_gen_pv
            self._update_Qmax_Qmin_Pmax_Pmin_Pgen(P_each_gen_pv,Pmax,Vsch_pv,Pmin)
            self.Set_up_for_BESS()
            self._run_power_flow()
            psspy.save(path_bessd)


            ##BESS C
            ierr = psspy.psseinit()
            ierr = psspy.case(path_bessc)
            self.Set_up_for_PV()
            Pmin = [0] * len(P_each_gen_pv)
            Pmax = P_each_gen_pv
            self._update_Qmax_Qmin_Pmax_Pmin_Pgen(P_each_gen_pv,Pmax,Vsch_pv,Pmin)
            self.Set_up_for_BESS()
            self._run_power_flow()
            psspy.save(path_bessc)

            ##PV+BESS
            # path = self.folder_for_results('basic_case_PVBESSD')
            # ierr = psspy.psseinit()
            # ierr = psspy.case(path)            
        except Exception as e:
            return str(e)

        return [res1,res2]


    def BESS_alone(self,Pnet,Qnet):
        if Pnet <= 0:
            ## if pnet <0 run BESS D and C else run BESS C
            ierr = psspy.psseinit()
            ierr = psspy.case(self.file)
            res = self.tune_BESS_alone(Pnet,Qnet)
            if type(res) is str:
                return f"Tune BESSD Fail!!!!\n{res}"
            else:
                P_each_gen = res[0]
                Vsch = res[1]
                res1 = self.tune_BESS_alone(-Pnet,Qnet)
                if type(res1) is str:
                    return f"Tune BESSC Fail!!!!\n{res1}"
                else:
                    P_each_gen1 = res1[0]
                    Vsch1 = res1[1]
                    path = self.folder_for_results("basic_case_BessD_alone")
                    ierr = psspy.psseinit()
                    ierr = psspy.case(path)
                    self.Set_up_for_BESS()
                    Pmin = P_each_gen1
                    Pmax = P_each_gen
                    warn = self._update_Qmax_Qmin_Pmax_Pmin_Pgen(P_each_gen,Pmax,Vsch,Pmin)
                    if type(warn) is str:
                        return warn
                    self._run_power_flow()
                    psspy.save(path)

                    path1 = self.folder_for_results("basic_case_BessC_alone")
                    ierr = psspy.psseinit()
                    ierr = psspy.case(path1)
                    self.Set_up_for_BESS()
                    Pmin = P_each_gen1
                    Pmax = P_each_gen
                    warn = self._update_Qmax_Qmin_Pmax_Pmin_Pgen(P_each_gen1,Pmax,Vsch1,Pmin)
                    if type(warn) is str:
                        return warn
                    self._run_power_flow()
                    psspy.save(path1)
                    res = [P_each_gen,Vsch,path,P_each_gen1,Vsch1,path1]
                    return res
        else:
            return "Tune BESS Pnet <0!!!!"
    def PV_alone(self,Pnet,Qnet):
        if Pnet >=0 :
            return "TUNE PV PNET < 0!!!!!!!!"
        ierr = psspy.psseinit()
        ierr = psspy.case(self.file)
        res_p_v=self._tune_PV_alone(Pnet,Qnet)
        if type(res_p_v) is str:
            return res_p_v
        else:
            P_each_gen = res_p_v[0]
            Vsch = res_p_v[1]
            path = self.folder_for_results("basic_case_PV_alone")
            ierr = psspy.psseinit()
            ierr = psspy.case(path)
            self.Set_up_for_PV()
            Pmin = [0] * len(P_each_gen)
            Pmax = P_each_gen
            warn = self._update_Qmax_Qmin_Pmax_Pmin_Pgen(P_each_gen,Pmax,Vsch,Pmin)
            if type(warn) is str:
                return warn
            self._run_power_flow()
            psspy.save(path)
            res = [P_each_gen,Vsch,path]
            return res

    def tune_BESS_alone(self,Pnet,Qnet):
        self.Set_up_for_BESS()
        res = self.Tune_P_and_V(Pnet,Qnet)
        return  res

    def _tune_PV_alone(self,Pnet,Qnet):
        self.Set_up_for_PV()
        res = self.Tune_P_and_V(Pnet,Qnet)
        return res

    def Tune_P_and_V(self,Pnet,Qnet):
        gen_number,poi_number,id_gen,id_poi =  self.count_gen()
        quantity_gen = self._get_gen_numbers(gen_number,id_gen)
        sum_gen = sum(quantity_gen)
        ## prepare
        Vsch = 1.0
        self.prepare_setup_P_Q_9999(gen_number,id_gen)
        self._update_voltage_schedule(gen_number, 1)
        ##start
        while True:
            ## P
            P_each_gen = self._adjust_active_power(Pnet)
            if type(P_each_gen) is str :
                return P_each_gen

            ## check đã chạy cả power flow
            Ppoi = self._get_power_at_poi(poi_number[0],id_poi[0])
            Qpoi = self._get_reactive_power_at_poi(poi_number[0],id_poi[0])
            if abs(Ppoi - Pnet) <= self.ERROR_TOLERANCE and abs(Qpoi-Qnet) <= self.ERROR_TOLERANCE1:
                print(P_each_gen)
                return [P_each_gen,Vsch]

            Vsch = self._adjust_reactive_power(Qnet)
            if type(Vsch) is str :
                return Vsch
            ## check again
            Ppoi = self._get_power_at_poi(poi_number[0],id_poi[0])
            Qpoi = self._get_reactive_power_at_poi(poi_number[0],id_poi[0])
            if abs(Ppoi - Pnet) <= self.ERROR_TOLERANCE and abs(Qpoi-Qnet) <= self.ERROR_TOLERANCE1:
                return [P_each_gen, Vsch]
    def LEAD_LAG(self,Pnet):
        Pnet = abs(Pnet)
        pf = 0.95
        Qnet = math.sqrt((Pnet/pf)**2-Pnet**2)
        ierr = psspy.psseinit()
        ierr = psspy.case(self.file)
        gen_number,poi_number,id_gen,id_poi = self.count_gen()
        ## folder result
        new_folder = os.path.join(self.folder, 'LEADLAG')
        if os.path.exists(new_folder):
            shutil.rmtree(new_folder)
        os.makedirs(new_folder, exist_ok=True)
        ## set up for lag
        Vsch_lag = self._adjust_reactive_power(-Qnet)

        warn = self.off_shunt()
        Vsch_lead =self._adjust_reactive_power(Qnet)

        if type(Vsch_lag) is not str:
            hv_ = os.path.join(new_folder, os.path.basename(self.file))
            hv_path = os.path.join(new_folder, "project_lag_hv.sav" )
            shutil.copy(self.file, new_folder) 
            os.rename(hv_,hv_path)
            ierr = psspy.psseinit()
            ierr = psspy.case(hv_path)        
            self._update_voltage_schedule(gen_number,Vsch_lag)
            self._run_power_flow()
            psspy.save(hv_path)

            lv_path = os.path.join(new_folder, "project_lag_lv.sav")
            shutil.copy(hv_path, lv_path)


        if type(Vsch_lead) is not str:
            hv1_ = os.path.join(new_folder, os.path.basename(self.file))
            hv_path1 = os.path.join(new_folder, "project_lead_hv.sav" )
            shutil.copy(self.file, new_folder)
            os.rename(hv1_,hv_path1)
            ierr = psspy.psseinit()
            ierr = psspy.case(hv_path1) 
            warn = self.off_shunt()  
            if type(warn) is str:
                Vsch_lead ='SHUNT OFF FAIL'    
            self._update_voltage_schedule(gen_number,Vsch_lead)
            self._run_power_flow()
            psspy.save(hv_path1)

            lv_path1 = os.path.join(new_folder, "project_lead_lv.sav")
            shutil.copy(hv_path1, lv_path1)

        return Vsch_lag,Vsch_lead

    def auto_Lead_Lag(self,Pnet):
        try:
            res  = self._tune_Vsch_with_33tap(Pnet)
            return res 
        except Exception as e:
            warn = f"Vsch 33 tap\n{str(e)}"
            return warn

    def _tune_Vsch_with_33tap(self,Pnet):
        Pnet = abs(Pnet)
        pf = 0.95
        Qnet = math.sqrt((Pnet/pf)**2-Pnet**2)
        ierr = psspy.psseinit()
        ierr = psspy.case(self.file)
        gen_number,poi_number,id_gen,id_poi = self.count_gen()

        res = self._get_tap_ratio_()
        if type(res) is str:
            return res
        
        step = res[0]
        tap_position = res[1]
        r_min = res[2]
        res ={}
        res['LAG'] = None
        res['LEAD'] = None
        res['auto_LAG'] = None
        res['auto_LEAD'] = None
        
        res_lag = {}
        res_lead = {}
        res_lag_auto = {}
        res_lead_auto = {}
        ierr, bus = psspy.atr3int(sid = -1 , flag=1, entry = 1, string = ['WIND1NUMBER','WIND2NUMBER','WIND3NUMBER'] )
        ierr, ID  = psspy.atr3char(sid = -1 , flag=1, entry = 1, string = ['ID'])

        for i in range(tap_position):
            ## flag =2 for all XFMR 
            ratio = r_min + i*step
            if ierr ==0:
                self._off_auto_tap_update_ratio(bus,ID,ratio)
                Vsch_lag = self._adjust_reactive_power(-Qnet)
                if type(Vsch_lag) is not str:
                    res_lag[round(ratio,5)] = round(Vsch_lag,6)
                ##auto
                    self._update_voltage_schedule(gen_number,Vsch_lag)
                    self._on_auto_tap_update_ratio(bus,ID,ratio)
                    self._run_power_flow()
                    ierr, ratio1 = psspy.awndreal(sid = -1, owner = 1 , ties =3 , flag=1, entry =2 , string = ['RATIO'])    
                    if abs(ratio1[0][0]-ratio) <0.00001:
                        res_lag_auto[round(ratio,5)] = round(Vsch_lag,6)

        for i in range(tap_position):
            ## flag =2 for all XFMR 
            print(ID)
            print(bus)
            ratio = r_min + i*step
            if ierr ==0:
                self._off_auto_tap_update_ratio(bus,ID,ratio)
                self.off_shunt()
                Vsch_lead = self._adjust_reactive_power(Qnet)
                if type(Vsch_lead) is not str:
                    res_lead[round(ratio,5)] = round(Vsch_lead,6)

                    self._update_voltage_schedule(gen_number,Vsch_lead)
                    self._on_auto_tap_update_ratio(bus,ID,ratio)
                    
                    self._run_power_flow()
                    ierr, ratio1 = psspy.awndreal(sid = -1, owner = 1 , ties =3 , flag=1, entry =2 , string = ['RATIO'])    
                    if abs(ratio1[0][0]-ratio) <0.00001:
                        res_lead_auto[round(ratio,5)] = round(Vsch_lead,6)

        res['LAG'] = res_lag
        res['LEAD'] = res_lead
        res['auto_LAG'] = res_lag_auto
        res['auto_LEAD'] = res_lead_auto
        df_list = []
        for key, values in res.items():
            for k, v in values.items():
                df_list.append({"Type": key, "Ratio": k, "Vsch": v})
        df = pd.DataFrame(df_list)
        # Ghi vào Excel
        try:
            output_path = self.folder + "\lead_lag_33tap.xlsx"
            df.to_excel(output_path, index=False, sheet_name="Lead_Lag")
        except:
            return "Close Excel Please"
        print(res)
        return res 

    def _off_auto_tap_update_ratio(self,bus,ID,ratio):
        for j in range(len(bus[0])):
            psspy.three_wnd_winding_data_5(bus[0][j],bus[1][j],bus[2][j],ID[0][j],1,[_i,_i,_i,_i,_i,-1],[ratio,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    def _on_auto_tap_update_ratio(self,bus,ID,ratio):
        for k in range(len(bus[0])):
            psspy.three_wnd_winding_data_5(bus[0][k],bus[1][k],bus[2][k],ID[0][k],1,[_i,_i,_i,_i,_i,1],[ratio,_f,_f,_f,_f,_f,_f,_f,_f,_f],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f])
    
    def _get_tap_ratio_(self):
        try:
            ierr, step_size_rmin = psspy.awndreal(sid = -1 , flag = 1 , entry = 1 , string = ['STEP','RMIN'])
            ierr, tap_position = psspy.awndint(sid = -1 , flag = 1 , entry = 1 , string = 'NTPOSN')

            step = round(step_size_rmin[0][0],5)
            r_min = round(step_size_rmin[1][0],5)
            tap_position = round(tap_position[0][1])
            return [step,tap_position,r_min]
        except Exception as e:
            print("Get Data 3XFMR Fail")   
            warn =  f'Get Data 3XFMR Fail\n{str(e)}'
            return warn
    def off_shunt(self):
        try:
            ierr, bus_number = psspy.aswshint(sid =-1, flag = 1, string = 'NUMBER')
            ierr, bus_ID = psspy.aswshchar(sid =-1, flag = 1, string = 'ID')
            for i,bus in enumerate(bus_number[0]):
                psspy.switched_shunt_chng_5(bus,bus_ID[0][i],[_i,_i,_i,_i,_i,_i,_i,_i,_i,_i,0,0,_i,1,1,1,1,1,1,1,_i],[_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],_s)
        except:
            return "Off Shunt Fail"
        
    
    def count_gen(self):
        ### cần update tự lấy id PJM có trường hợp k lấy được -> có thể lấy theo machine thay vì bus
        gen_number=[]
        poi_number=[]
        id_gen = []
        id_poi =[]
        ierr, bus_numbers = psspy.abusint(-1, 2, 'NUMBER')  # Get all bus numbers
        ierr, bus_types = psspy.abusint(-1, 2, 'TYPE')      # Get all bus type codes

        for i in range(len(bus_numbers[0])):
            if bus_types[0][i]==2:
                gen_number.append(bus_numbers[0][i])
            if bus_types[0][i]==3:
                poi_number.append(bus_numbers[0][i])

        ierr, bus_machine = psspy.amachint(-1, 1, 'NUMBER')
        ierr, id_machine = psspy.amachchar(-1, 1, 'ID')

        for i,bus_number in enumerate(gen_number):
            for j,bus_machine1 in enumerate(bus_machine[0]):
                if bus_number == bus_machine1:
                    id_gen.append(id_machine[0][j])

        for i,bus_machine1 in enumerate(bus_machine[0]):
            if bus_machine1 == poi_number[0]:
                id_poi.append(id_machine[0][i])

        return gen_number,poi_number,id_gen,id_poi
    def Set_up_for_BESS(self):
        ## turn off bess and on PV
        ## if PV =1 thì all là PV
        ## on BESS

        try:
            branchbess1 = list(map(int, self.branchbess.split(",")))
            branchpv1 = list(map(int, self.branchpv.split(",")))
            for bus in branchbess1:
                psspy.recn(bus)
            ## Turn off PV
            for bus in branchpv1:
                psspy.dscn(bus)
            return False
        except:
            return True
    def Set_up_for_PV(self):
        try:
            branchbess1 = list(map(int, self.branchbess.split(",")))
            branchpv1 = list(map(int, self.branchpv.split(",")))
            for bus in branchpv1:
                psspy.recn(bus)
            ## Turn off bess
            for bus in branchbess1:
                psspy.dscn(bus)
            return False
        except:
            return True
    def Set_up_for_PVBESS(self):
        try:
            branchbess1 = list(map(int, self.branchbess.split(",")))
            branchpv1 = list(map(int, self.branchpv.split(",")))
            for bus in branchbess1:
                psspy.recn(bus)
            ## Turn off PV
            for bus in branchpv1:
                psspy.recn(bus)
            return False
        except:
            return True
    def folder_for_results(self,name):
        new_folder = os.path.join(self.folder, name)
        if os.path.exists(new_folder):
            shutil.rmtree(new_folder)
        os.makedirs(new_folder, exist_ok=True)

        new_path_file = os.path.join(new_folder, os.path.basename(self.file))
        shutil.copy(self.file, new_folder)
        return new_path_file
    def _adjust_active_power(self,Pnet):
        ## lấy Pnet tính toán Pgen cần thiết (Pgen chia cho số nhánh)
        ## chắc chắn sẽ có tổn thất cần chọn chiều để tool
        ## cần giới hạn 7 hoặc 6 chữ số để dừng lại
        ## input Pnet = ?
        Pgen = - Pnet

        gen_number,poi_number,id_gen,id_poi =  self.count_gen()

        ## đếm số lượng gen
        quantity_gen = self._get_gen_numbers(gen_number,id_gen)
        ## tổng gen để lấy tỉ lệ scale gen
        sum_gen = sum(quantity_gen)

        MIN_DELTA = 10^-6
        for i in range(1,self.MAX_ITER):
            delta = 10 / (10 ** i)

            while True:
                ## pgen1 công suất từng gen
                P_each_gen = self._calculate_generator_power(Pgen, quantity_gen, sum_gen)
                ## update gen
                if  self._update_generator_power(gen_number,id_gen, P_each_gen):
                ## nếu cập nhật gen có lỗi thì break
                    print('update gen false')
                    return 'Update Gen False!!!!!!!'

                if self._run_power_flow():
                    print('run power flow false')
                    return 'Run Power Flow False'
                Ppoi = self._get_power_at_poi(poi_number[0],id_poi[0])

                ## CHECK
                if abs(Ppoi - Pnet) <= self.ERROR_TOLERANCE:
                    return  P_each_gen

                ## Ppoi = -80 > Pnet = -90 đang thiếu công suất Pgen cần thêm công suất với trường hợp Ppoi =-90 tool thuận
                ## Ppoi = 80 < Pnet = 90 Pgen đang nhận ít công suất cần tăng theo chiều âm tool ngược
                ## need to test with 2 case positive and negative at poi
                ## delta start = 1
                if (Ppoi - Pnet) > 0:
                    Pgen += delta
                else:
                    Pgen -= delta
                    break

                ## check xem delta đã có bao nhiêu chữ số
                # if delta < MIN_DELTA:
                #     return None,None
        return 'DELTA P IS TOO SMALL'
    def ____adjust_reactive_power_backup(self):
        ## get Qpoi for first search
        ### need to set up binary research chwua giưới hạn 1.1 và 0.9
        gen_number,poi_number,id_gen,id_poi =  self.count_gen()

        Vsch = 1
        self._update_voltage_schedule(gen_number, 1)
        if self._run_power_flow():
            return 'Run Power Flow False'
        Q_first = self._get_reactive_power_at_poi(poi_number[0],id_poi[0])
        for decimal_places in range(2, self.MAX_ITER):
            delta = 1 / (10 ** decimal_places)
            while True:
                if Q_first < 0 :
                    Vsch -= delta
                else:
                    Vsch += delta
                if self._update_voltage_schedule(gen_number, Vsch):
                    print("update Vsch false")
                    return 'Update Vsch False'
                if self._run_power_flow():
                    print("run power flow flase")
                    return 'Run Power Flow False'
                Qpoi = self._get_reactive_power_at_poi(poi_number[0],id_poi[0])

                if abs(Qpoi) <= self.ERROR_TOLERANCE1:
                    return Vsch

                if Q_first * Qpoi <= 0:
                    if Q_first < 0:
                        Vsch += delta
                    else:
                        Vsch -= delta
                    break

                # if delta < MIN_DELTA:
                #     return None
        return 'DELTA Q IS TOO LARGE'
    def _adjust_reactive_power(self,Qnet):
        ## get Qpoi for first search
        ### need to set up binary research chwua giưới hạn 1.1 và 0.9
        gen_number,poi_number,id_gen,id_poi =  self.count_gen()

        Vsch = 1
        self._update_voltage_schedule(gen_number, 1)
        if self._run_power_flow():
            return 'Run Power Flow False'
        Q_first = self._get_reactive_power_at_poi(poi_number[0],id_poi[0])
        for decimal_places in range(2, self.MAX_ITER):
            delta = 1 / (10 ** decimal_places)
            while True:
                if Q_first - Qnet < 0 :
                    Vsch -= delta
                    flag = True
                else:
                    Vsch += delta
                    flag = False
                if self._update_voltage_schedule(gen_number, Vsch):
                    print("update Vsch false")
                    return 'Update Vsch False'
                if self._run_power_flow():
                    print("run power flow flase")
                    return 'Run Power Flow False'
                Qpoi = self._get_reactive_power_at_poi(poi_number[0],id_poi[0])

                if abs(Qpoi-Qnet) <= 0.003:
                    return Vsch
                if Vsch > 1.1 or Vsch <0.9:
                    return "Vsch out of range"
                if flag == True and (Qpoi-Qnet) >0 :
                    Vsch += delta
                    break
                elif flag == False and (Qpoi-Qnet)<0 :
                    Vsch -= delta
                    break

                # if delta < MIN_DELTA:
                #     return None
        return 'DELTA Q IS TOO SMALL'
    def _calculate_generator_power(self,Pgen, quantity_gen, sum_gen):
        return [Pgen / sum_gen * num for num in quantity_gen]
    def _update_generator_power(self,gen_number, id_gen,P_each_gen):
        flag = False
        for i, gen in enumerate(gen_number):
            ierr = psspy.machine_chng_4(gen, id_gen[i], [_i]*7, [P_each_gen[i]] + [_f]*16, "")
            if ierr != 0:
                flag = True
        return flag
    def _run_power_flow(self):
        flag = False
        ierr = psspy.fnsl([1,1,0,0,1,1,0,0])
        if ierr != 0:
            flag = True
        return flag
    def _get_gen_numbers(self,gen_number,id_gen):
        if len(gen_number) > 1:
            if self.ScaleGen == 'No':
                return [1 for _ in gen_number]
            else:
                quantity_gen =[]
                for i,pgen in enumerate(gen_number):
                    ierr, MVAbase = psspy.macdat(gen_number[i], id_gen[i], 'MBASE')
                    quantity_gen.append(MVAbase)
            return quantity_gen  # Simplified: always return 1 for each generator
        else:
            return [1]
    def _get_power_at_poi(self,poi_number,id_poi):
        ierr, Ppoi = psspy.macdat(poi_number, id_poi, 'P')
        return Ppoi
    def _get_reactive_power_at_poi(self,poi_number,id_poi) :
        ierr, Qpoi = psspy.macdat(poi_number, id_poi, 'Q')
        return Qpoi
    def _update_voltage_schedule(self, gen_number, Vsch):
        flag = False
        for gen in gen_number:
            ## need to check xem có phải thay đổi tại plant hay bus
            ierr  = psspy.plant_chng_4(gen, 0, [_i, _i], [Vsch, _f])
            if ierr != 0:
                flag = True
        return flag
    def _update_Qmax_Qmin_Pmax_Pmin_Pgen(self,P_each_gen,Pmax,Vsch,Pmin):
        try:
            gen_number,poi_number,id_gen,id_poi =  self.count_gen()
            for i,pgen in enumerate(P_each_gen):
                ierr, MVAbase = psspy.macdat(gen_number[i], id_gen[i], 'MBASE')
                Qmax = math.sqrt(MVAbase**2 - Pmax[i]**2)
                if self.PQCurve == 'No':
                    psspy.machine_chng_4(gen_number[i],id_gen[i],[_i,_i,_i,_i,_i,_i,_i],[pgen,_f,Qmax,-Qmax,Pmax[i],Pmin[i],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],"")
                else:
                    psspy.machine_chng_4(gen_number[i],id_gen[i],[_i,_i,_i,_i,_i,_i,_i],[pgen,_f,9999,-9999,Pmax[i],Pmin[i],_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],"")
                psspy.plant_chng_4(gen_number[i],0,[_i,_i],[Vsch,_f])
            return True
        except:
            return "Update P and Q fail maybe Pgen > MVABase"
    def prepare_setup_P_Q_9999(self,gen_number,id_gen):
        for i,gen_num in enumerate(gen_number):
            psspy.machine_chng_4(gen_number[i],id_gen[i],[_i,_i,_i,_i,_i,_i,_i],[_f,_f,9999,-9999,9999,-9999,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f,_f],"")
class Report_MQT:
    ## this function help get MPT ratio, Vsch Gen,
    ##  chuyển file để viết report
    def __init__(self,ui):
        self.ui = ui
    def prepare_folder_for_report(self):
        folder = self.ui.lineINPUT.text()
        if  not folder:
            self.ui.msg_box("folder")
            return None,None
        else:

            try:
                ### creat folder report MQT out of DMV
                path_reportMQT = os.path.dirname(folder)
                path_reportMQT = os.path.join(path_reportMQT,"ReportMQT")
                if os.path.exists(path_reportMQT):
                    shutil.rmtree(path_reportMQT)
                os.makedirs(path_reportMQT)

                path = os.path.join(folder, "RESULTs")
                return path_reportMQT,path
            except:
                self.ui.resultsTextEdit.append("FAILT WHEN CREAT REPORT FILE !!!")
                return None, None
    def remove(self,path,subfolder,a):
        dest_subfolder_path = os.path.join(path, subfolder.name)  # Đường dẫn đích sao chép
        # Sao chép thư mục con vào thư mục đích
        renamed_subfolder_path = os.path.join(path, subfolder.name+a)
        try:

            shutil.copytree(subfolder.path, dest_subfolder_path)
            os.rename(dest_subfolder_path, renamed_subfolder_path)
        except Exception as e:
            self.ui.resultsTextEdit.append(f" {subfolder.name}: {e}")
    def move_file(self):
        path_reportMQT,path_results = self.prepare_folder_for_report()

        for folder in os.scandir(path_results):
            if folder.is_dir():

                if folder.name == 'PROJECT':
                    project_path = os.path.join(path_results, "PROJECT")

                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path  # Đường dẫn thư mục con trong 'PROJECT'
                            if subfolder.name == 'TEST1_FS_FS':
                                self.remove(path_reportMQT, subfolder,"d")
                            if subfolder.name == 'TEST2_VOLTDOWN_VOLT':
                                self.remove(path_reportMQT, subfolder,"d")
                            if subfolder.name == 'TEST3_VOLTUP_VOLT':
                                self.remove(path_reportMQT, subfolder,"d")
                            if subfolder.name == 'TEST4_FRQDOWN_FREQ': ##nhr
                                self.remove(path_reportMQT, subfolder,"nohrd")
                            if subfolder.name == 'TEST8_SCR2_SCR2':
                                self.remove(path_reportMQT, subfolder,"d")


                if folder.name == 'PROJECT_80':
                    project_path = os.path.join(path_results, "PROJECT_80")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST4_FRQDOWN_FREQ':
                                self.remove(path_reportMQT, subfolder,"d")
                            if subfolder.name =='TEST5_FRQUP_FREQ':
                                self.remove(path_reportMQT, subfolder,"d")


                if folder.name == 'PROJECT_LAG_HV':
                    project_path = os.path.join(path_results, "PROJECT_LAG_HV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGd")


                if folder.name == 'PROJECT_LAG_LV':
                    project_path = os.path.join(path_results, "PROJECT_LAG_LV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGd")

                if folder.name == 'PROJECT_LEAD_HV':
                    project_path = os.path.join(path_results, "PROJECT_LEAD_HV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADd")


                if folder.name == 'PROJECT_LEAD_LV':
                    project_path = os.path.join(path_results, "PROJECT_LEAD_LV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADd")


                if folder.name == '2PROJECT_LAG_HV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "2PROJECT_LAG_HV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGd_PREFER")

                if folder.name == '2PROJECT_LAG_LV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "2PROJECT_LAG_LV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGd_PREFER")

                if folder.name == '2PROJECT_LEAD_HV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "2PROJECT_LEAD_HV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADd_PREFER")

                if folder.name == '2PROJECT_LEAD_LV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "2PROJECT_LEAD_LV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADd_PREFER")




        ### need to check
                if folder.name == 'CPROJECT':
                    project_path = os.path.join(path_results, "CPROJECT")

                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path  # Đường dẫn thư mục con trong 'PROJECT'
                            if subfolder.name == 'TEST1_FS_FS':
                                self.remove(path_reportMQT, subfolder,"c")
                            if subfolder.name == 'TEST2_VOLTDOWN_VOLT':
                                self.remove(path_reportMQT, subfolder,"c")
                            if subfolder.name == 'TEST3_VOLTUP_VOLT':
                                self.remove(path_reportMQT, subfolder,"c")
                            if subfolder.name == 'TEST5_FRQUP_FREQ': ##nhr
                                self.remove(path_reportMQT, subfolder,"nohrc")
                            if subfolder.name == 'TEST8_SCR2_SCR2':
                                self.remove(path_reportMQT, subfolder,"c")


                if folder.name == 'CPROJECT_80':
                    project_path = os.path.join(path_results, "CPROJECT_80")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST4_FRQDOWN_FREQ':
                                self.remove(path_reportMQT, subfolder,"c")
                            if subfolder.name =='TEST5_FRQUP_FREQ':
                                self.remove(path_reportMQT, subfolder,"c")



                if folder.name == 'CPROJECT_LAG_HV':
                    project_path = os.path.join(path_results, "CPROJECT_LAG_HV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGc")

                if folder.name == 'CPROJECT_LAG_LV':
                    project_path = os.path.join(path_results, "CPROJECT_LAG_LV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGc")

                if folder.name == 'CPROJECT_LEAD_HV':
                    project_path = os.path.join(path_results, "CPROJECT_LEAD_HV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADc")

                if folder.name == 'CPROJECT_LEAD_LV':
                    project_path = os.path.join(path_results, "CPROJECT_LEAD_LV")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADc")

                if folder.name == '1CPROJECT_LAG_HV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "1CPROJECT_LAG_HV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGc_PREFER")

                if folder.name == '1CPROJECT_LAG_LV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "1CPROJECT_LAG_LV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LAGc_PREFER")

                if folder.name == '1CPROJECT_LEAD_HV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "1CPROJECT_LEAD_HV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST6_HVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADc_PREFER")

                if folder.name == '1CPROJECT_LEAD_LV_PREFER_PVBESS2':
                    project_path = os.path.join(path_results, "1CPROJECT_LEAD_LV_PREFER_PVBESS2")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path
                            if subfolder.name =='TEST7_LVRT_VOLT':
                                self.remove(path_reportMQT, subfolder,"_LEADc_PREFER")


        for folder in os.scandir(path_results):
            if folder.is_dir():

                if folder.name == 'PROJECT_VOLT':
                    project_path = os.path.join(path_results, "PROJECT_VOLT")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path  # Đường dẫn thư mục con trong 'PROJECT'
                            if subfolder.name == 'TEST2_VOLTDOWN_VOLT':
                                renamed_subfolder_path = os.path.join(path_reportMQT, subfolder.name+'d')
                                if os.path.exists(renamed_subfolder_path):
                                    shutil.rmtree(renamed_subfolder_path)
                                self.remove(path_reportMQT, subfolder,"d")
                            if subfolder.name == 'TEST3_VOLTUP_VOLT':
                                renamed_subfolder_path = os.path.join(path_reportMQT, subfolder.name+'d')
                                if os.path.exists(renamed_subfolder_path):
                                    shutil.rmtree(renamed_subfolder_path)
                                self.remove(path_reportMQT, subfolder,"d")

                if folder.name == 'CPROJECT_VOLT':
                    project_path = os.path.join(path_results, "CPROJECT_VOLT")
                    for subfolder in os.scandir(project_path):
                        if subfolder.is_dir():
                            subfolder_path = subfolder.path  # Đường dẫn thư mục con trong 'PROJECT'
                            if subfolder.name == 'TEST2_VOLTDOWN_VOLT':
                                renamed_subfolder_path = os.path.join(path_reportMQT, subfolder.name+'c')
                                if os.path.exists(renamed_subfolder_path):
                                    shutil.rmtree(renamed_subfolder_path)

                                self.remove(path_reportMQT, subfolder,"c")
                            if subfolder.name == 'TEST3_VOLTUP_VOLT':
                                renamed_subfolder_path = os.path.join(path_reportMQT, subfolder.name+'c')
                                if os.path.exists(renamed_subfolder_path):
                                    shutil.rmtree(renamed_subfolder_path)
                                self.remove(path_reportMQT, subfolder,"c")

        self.ui.resultsTextEdit.append("CREATED FOLDER FOR REPORT DONE!!!\n")
        return path_reportMQT
    def get_ratio_and_vsch(self,path):
        res = {}
        print(path)
        try:
            for folder in os.listdir(path):
                sub_path = os.path.join(path,folder)
                for file in os.listdir(sub_path):
                    if file.endswith(".sav") and not file.endswith("cnv.sav") :
                        path_sav = os.path.join(sub_path, file)
                        ierr = psspy.psseinit()
                        ierr = psspy.case(path_sav)


                        ierr, ratio = psspy.awndreal(sid = -1, owner = 1 , ties =3 , flag=1, entry =2 , string = ['RATIO'])
                        ###gen bus
                        ierr, type_code = psspy.agenbusint(sid= -1 , flag =1 , string = 'TYPE')
                        print(type_code)
                        ierr, type_code = psspy.abusint(sid = -1,flag = 1,string = 'TYPE')
                        print(type_code)
                        ## Vsch bus
                        ierr, vsch = psspy.abusreal(sid =-1, flag =1, string = 'PU')
                        print(vsch)
                        ierr, capbank = psspy.aswshint(sid=-1, flag =1, string='STATUS')
                        print(capbank)
                        for i in range(len(type_code[0])):
                            flag= False
                            vsch_bus = None
                            if type_code[0][i] == 2 :
                                vsch_bus = round(vsch[0][i],6)
                                flag = True
                                break
                            if flag:
                                break
                        flag_capbank = 'off'
                        for i in range(len(capbank[0])):
                            if capbank[0][i] == 1 :
                                flag_capbank = 'on'

                        if vsch_bus is None:
                            vsch_bus = 0

                        if ratio[0]:
                            ratio_1 = round(ratio[0][0],6)
                        else:
                            ratio_1 = 0

                        key = os.path.basename(os.path.dirname(path_sav))
                        res[key] =[ratio_1,vsch_bus,flag_capbank]
                        print(res)
            for key,value in res.items():
                line = f"{key}:Ratio:{value[0]},Vsch:{value[1]},CapBank:{value[2]}\n "
                self.ui.resultsTextEdit.append(line)
            return res
        except Exception as e:
            self.ui.resultsTextEdit.append(f" {subfolder.name}: {e}")
    def main(self):
        ## this function when after get the folder report then get mpt and vsch
        try:
            path_reportMQT = self.move_file()
            res = self.get_ratio_and_vsch(path_reportMQT)
            path, ok = QInputDialog.getText(
                None,
                "Excel Report",       # Tiêu đề hộp thoại
                "Path Excel Report:"  # Nội dung hướng dẫn
            )

            if ok and path:  # Nếu người dùng bấm OK và có nhập giá trị

                self.report_excel(path,res)
                self.ui.resultsTextEdit.append(path)
                self.ui.resultsTextEdit.append("Report Excel Done!!")
                # Tiếp tục xử lý với đường dẫn `text`
        except Exception as e:
            fail = str(e)
            return self.ui.resultsTextEdit.append(f"SOMETHING WENT WRONG WHEN CREAT FOLDER REPORT MQT AND GET RATIO,VSCH!!!{fail}")
    def input_value(self,res,key,sheet,value):
        sheet[f'G{value}'] = key
        sheet[f'H{value}'] = res[key][0]
        sheet[f'I{value}'] = res[key][1]
        sheet[f'J{value}'] = res[key][2]
        sheet[f'B{value}'] = 'Yes'
    def report_excel(self,path,res):
        wb = openpyxl.load_workbook(path)
        sheet = wb['ControlNarrative']
        for key in res :
            if key == 'TEST1_FS_FSd':
                self.input_value(res,key,sheet,'3')
            if key =='TEST2_VOLTDOWN_VOLTd':
                self.input_value(res,key,sheet,'4')
            if key =='TEST3_VOLTUP_VOLTd':
                self.input_value(res,key,sheet,'5')
            if key =='TEST4_FRQDOWN_FREQnohrd':
                self.input_value(res,key,sheet,'6')
            if key =='TEST4_FRQDOWN_FREQd':
                self.input_value(res,key,sheet,'7')
            if key =='TEST5_FRQUP_FREQd':
                self.input_value(res,key,sheet,'8')
            if key =='TEST6_HVRT_VOLT_LEADd':
                self.input_value(res,key,sheet,'9')
            if key =='TEST6_HVRT_VOLT_LAGd':
                self.input_value(res,key,sheet,'10')
            if key =='TEST7_LVRT_VOLT_LEADd':
                self.input_value(res,key,sheet,'11')
            if key =='TEST7_LVRT_VOLT_LAGd':
                self.input_value(res,key,sheet,'12')
            if key =='TEST6_HVRT_VOLT_LEADd_PREFER':
                self.input_value(res,key,sheet,'15')
            if key =='TEST6_HVRT_VOLT_LAGd_PREFER':
                self.input_value(res,key,sheet,'16')
            if key =='TEST7_LVRT_VOLT_LEADd_PREFER':
                self.input_value(res,key,sheet,'17')
            if key =='TEST7_LVRT_VOLT_LAGd_PREFER':
                self.input_value(res,key,sheet,'18')
            if key =='TEST8_SCR2_SCR2d':
                self.input_value(res,key,sheet,'21')

            if key =='TEST1_FS_FSc':
                self.input_value(res,key,sheet,'22')
            if key =='TEST2_VOLTDOWN_VOLTc':
                self.input_value(res,key,sheet,'23')
            if key =='TEST3_VOLTUP_VOLTc':
                self.input_value(res,key,sheet,'24')
            if key =='TEST5_FRQUP_FREQnohrc':
                self.input_value(res,key,sheet,'25')
            if key =='TEST5_FRQUP_FREQc':
                self.input_value(res,key,sheet,'26')
            if key =='TEST4_FRQDOWN_FREQc':
                self.input_value(res,key,sheet,'27')
            if key =='TEST6_HVRT_VOLT_LEADc':
                self.input_value(res,key,sheet,'28')
            if key =='TEST6_HVRT_VOLT_LAGc':
                self.input_value(res,key,sheet,'29')
            if key =='TEST7_LVRT_VOLT_LEADc':
                self.input_value(res,key,sheet,'30')
            if key =='TEST7_LVRT_VOLT_LEADc':
                self.input_value(res,key,sheet,'31')
            if key =='TEST6_HVRT_VOLT_LEADc_PREFER':
                self.input_value(res,key,sheet,'34')
            if key =='TEST6_HVRT_VOLT_LAGc_PREFER':
                self.input_value(res,key,sheet,'35')
            if key =='TEST7_LVRT_VOLT_LEADc_PREFER':
                self.input_value(res,key,sheet,'36')
            if key =='TEST7_LVRT_VOLT_LAGc_PREFER':
                self.input_value(res,key,sheet,'37')
            if key =='TEST8_SCR2_SCR2c':
                self.input_value(res,key,sheet,'40')
        wb.save(path)

if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
####    ID= ''
    license_dialog = LicenseCheckDialog()
    if license_dialog.check_license()  :

        # Hiển thị cửa sổ chính nếu hộp thoại kiểm tra license được chấp nhận
        MainWindow = QtWidgets.QMainWindow()
        ui = Ui_MainWindow()
        ui.setupUi(MainWindow)
        MainWindow.show()
        sys.exit(app.exec_())
    else:
        sys.exit()  # Thoát nếu hộp thoại license bị từ chối
