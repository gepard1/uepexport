import requests
import json
from PyQt5.QtWidgets import *
import sys
import pandas as pd
from lxml import html
import re

deps = ["101", "102", "103", "104", "105", "106", "107"]

type = {"Studia stacjonarne": "1", "Studia niestacjonarne": "2"}

studies = {
    "Ekonomii": "101",
    "Finansów i rynków finansowych": "102",
    "Gospodarki międzynarodowej": "107",
    "Informatyki i analiz ekonomicznych": "106",
    "Nauk o jakości": "105",
    "Rachunkowości i finansów przedsiebiorstw": "103",
    "Zarządzania": "103",
}

level = {"Studia pierwszego stopnia": "1", "Studia drugiego stopnia": "2"}

year = ["1", "2", "3"]

data = {"dep": "101", "cyc": "1", "year": "1", "type": "1"}

headers = {
    "Host": "app.ue.poznan.pl",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:99.0) Gecko/20100101 Firefox/99.0",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "pl,en-US;q=0.7,en;q=0.3",
    "Accept-Encoding": "gzip, deflate, br",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Content-Length": "27",
    "Origin": "https://app.ue.poznan.pl",
    "DNT": "1",
    "Connection": "keep-alive",
    "Referer": "https://app.ue.poznan.pl/Schedule/",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
}

url = "https://app.ue.poznan.pl/Schedule/Home/GetGroup"

group = "10111001"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.combo2 = QComboBox()
        self.combo3 = QComboBox()
        self.combo4 = QComboBox()
        self.combo5 = QComboBox()
        self.combo = QComboBox()

        win = QDialog()
        self.b1 = QPushButton(win)
        self.b1.setText("Eksportuj do csvpy")
        self.b1.clicked.connect(self.b1_clicked)

        for line in type:
            self.combo2.addItem(line)

        for line in studies:
            self.combo3.addItem(line)

        for line in level:
            self.combo4.addItem(line)

        for line in year:
            self.combo5.addItem(line)

        self.combo2.currentTextChanged.connect(self.text_changed2)
        self.combo3.currentTextChanged.connect(self.text_changed3)
        self.combo4.currentTextChanged.connect(self.text_changed4)
        self.combo5.currentTextChanged.connect(self.text_changed5)
        self.combo.currentIndexChanged.connect(self.text_changed)

        layout = QVBoxLayout()
        layout.addWidget(self.combo2)
        layout.addWidget(self.combo3)
        layout.addWidget(self.combo4)
        layout.addWidget(self.combo5)
        layout.addWidget(self.combo)
        layout.addWidget(self.b1)
        self.combo_edit()
        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)

    def text_changed2(self, s):
        self.type = type[s]
        data["type"] = self.type
        self.combo_edit()

    def text_changed3(self, s):
        self.studies = studies[s]
        data["dep"] = self.studies
        self.combo_edit()

    def text_changed4(self, s):
        self.level = level[s]
        data["cyc"] = self.level
        self.combo_edit()

    def text_changed5(self, s):
        self.year = s
        data["year"] = self.year
        self.combo_edit()

    def text_changed(self, index):
        x = requests.post(url=url, data=data, headers=headers)
        c = json.loads(x.text)
        self.group = c[index]["Value"]
        self.group_name = c[index]["Text"]

    def combo_edit(self):
        self.combo.clear()
        x = requests.post(url=url, data=data, headers=headers)
        c = json.loads(x.text)
        for group in c:
            self.combo.addItem(group["Text"])

    def b1_clicked(self):
        dep = data["dep"]
        cyc = data["cyc"]
        year = data["year"]
        type = data["type"]

        url = f"https://app.ue.poznan.pl/Schedule/Home/GetTimeTable?dep={dep}&cyc={cyc}&year={year}&group={self.group}&type={type}"

        page = requests.get(url)
        response = html.fromstring(page.content.decode("utf8"))

        rows = response.xpath("/html/body/table[1]/tbody")[0].findall("tr")

        row_list = []
        for T in rows:
            if len(T) == 3:
                row_list.append(T)

        plan_dict = {}

        for row in row_list:
            tds = row.findall("td")
            for td in tds:
                div = td.find("div")
                if len(div) > 1:
                    date = div[0].text_content()
                    match = re.search(r"\d{4}-\d{2}-\d{2}", date)
                    plan_dict[match.group()] = []
                    for i in div[1:]:
                        foo = list(filter(str.strip, i.text_content().split("\r\n")))
                        plan_dict[match.group()].append(foo)

        df_final = pd.DataFrame()
        for key in plan_dict:
            for lesson in plan_dict[key]:
                lesson_dict = {}
                lesson_dict["Subject"] = lesson[1]
                lesson_dict["Start date"] = key
                hours = lesson[0].split("-")
                lesson_dict["Start time"] = hours[0]
                lesson_dict["End time"] = re.search(r"\d{2}:\d{2}", hours[1]).group()
                lesson_dict["Description"] = lesson[3]
                lesson_dict["Location"] = lesson[2]
                df = pd.DataFrame([lesson_dict])
                df_final = pd.concat([df_final, df], ignore_index=True)

        df_final.to_csv(f"{self.group_name}.xlsx", encoding="utf-8")


app = QApplication(sys.argv)
w = MainWindow()
w.show()
app.exec_()
