#-*-encoding:utf8-*-

import requests
import time
import sys
import xlsxwriter
import multiprocessing
import queue
from bs4 import BeautifulSoup
from multiprocessing import Pool
from PyQt4.QtGui import *

import aladin_form

process_max = 8

'''
 *************************
 ****** 관련 데이터 *******
 *************************
'''

cid_list = {
    '1230' : '가정/요리/뷰티',
    '2551' : '만화',
    '1108' : '어린이',
    '656' : '인문학',
    '50246' : '초등참고서',
    '55890' : '건강/취미/레저',
    '4395' : '사전/기타',
    '13789' : '유아',
    '336' : '자기계발',
    '351' : '컴퓨터',
    '170' : '경제경영',
    '798' : '사회과학',
    '1196' : '여행',
    '1237' : '종교/역학',
    '2105' : '고전',
    '1' : '소설/시/희곡',
    '74' : '역사',
    '2030' : '좋은부모',
    '987' : '과학',
    '55889' : '에세이',
    '517' : '예술/대중문화',
    '1137' : '청소년',
    '8257' : '대학교재/전문서적',
    '1383' : '수험서/자격증',
    '1322' : '외국어',
    '50245' : '중/고등참고서'
}

offcode_list = {
    'sinsa' : '가로수길점',
    'gangnam' : '강남점',
    'geondae' : '건대점',
    'nowon' : '노원점',
    'daehakro' : '대학로점',
    'suyu' : '수유점',
    'sillim' : '신림점',
    'sinchon' : '신촌점',
    'yeonsinnae' : '연신내점',
    'jamsil' : '잠실롯데월드타워점',
    'sincheon' : '잠실신천점',
    'jongno' : '종로점',
    'hapjeong' : '합정점',
    'dongtan' : '동탄점',
    'bucheon' : '부천점',
    'buksuwon' : '북수원홈플러스점',
    'bundang' : '분당서현점',
    'yatap' : '분당야탑점',
    'sanbon' : '산본점',
    'suwon' : '수원점',
    'ilsan' : '일산점',
    'hwajeong' : '화정점',
    'sangmu' : '광주상무점',
    'gwangju' : '광주충장로점',
    'daegu' : '대구동성로점',
    'sangin' : '대구상인점',
    'daejeoncityhall' : '대전시청역점',
    'daejeon' : '대전은행점',
    'killy' : '부산경성대부경대역점',
    'deokcheon' : '부산덕천점',
    'dongbo' : '부산서면점',
    'centum' : '부산센텀점',
    'ulsan' : '울산점',
    'gyesan' : '인천계산홈플러스점',
    'jeonju' : '전주점',
    'cheonan' : '천안점',
    'cheongju' : '청주점',
    'guwol' : '인천구월점'
}



'''
 *************************
 **** 크롤링 관련 함수 ****
 *************************
'''

# 책에 대한 정보 갖고있는 객체
class Book:
    def __init__(self, _title, _isbn, _itemid, _stock, _price, _cid):
        self.title = _title
        self.isbn = _isbn
        self.itemid = _itemid
        self.stock = _stock
        self.price = _price # 중고 판매가 최저가
        self.cid = _cid # 카테고리 id
        self.fixprice = 0 # 정가 (절판이면 -1)
        self.myprice = 0
        self.isbn13 = ""

    def setMyPrice(self, factor_list):
        if self.fixprice == -1:
            self.myprice = self.price
            return

        sub_myprice = [self.price*factor_list[0],
                       self.fixprice*factor_list[1],
                       self.fixprice*factor_list[2],
                       factor_list[3]]

        # factor.1 적용하여 최대치 넘는다면
        if sub_myprice[0] >= sub_myprice[1]:
            self.myprice = sub_myprice[1]
        # factor.1 적용하여 최저치에 못 미친다면
        elif sub_myprice[0] <= sub_myprice[2]:
            self.myprice = sub_myprice[2]
        # factor.1 적용하여 최대~최저 사이라면
        else:
            self.myprice = sub_myprice[0]

        # 최저 보장가보다 낮은 경우
        if self.myprice < factor_list[3]:
            self.myprice = factor_list[3]

        self.myprice = int(self.myprice)

    def __repr__(self):
        return "< " + self.title + ", " + self.isbn + ", " + str(self.price) + ", " + str(self.fixprice) + ", " + str(self.stock) + " >\n"

# 카테고리 페이지 수
def getPages(offcode, cid):
    global book_list
    url = "http://used.aladin.co.kr/usedstore/wbrowse.aspx?offcode=" + offcode + "&cid=" + cid
    req = requests.get(url)
    bs = BeautifulSoup(req.text, 'html.parser')
    html = bs.find("div", {"id": "short"}).find("div", "numbox_last").find("a").get("href")

    count = range(1, int(html.split("'")[1])+1)
    offcodel = []
    cidl = []
    for i in count:
        offcodel.append(offcode)
        cidl.append(cid)
    return zip(offcodel, cidl, count)

# 카테고리 페이지 순회하면서 크롤링
def getBooks(args): #args[0] : offcode , args[1] : cid , args[2] : page

    if sys.getrecursionlimit() != 100000:
        sys.setrecursionlimit(100000)

    temp_list = []
    url = "http://off.aladin.co.kr/usedstore/wbrowse.aspx?offcode=" + args[0] + "&ItemType=0&BrowseTarget=AllView&ViewRowsCount=25&ViewType=Detail&PublishMonth=0&SortOrder=5&page=" + str(args[2]) + "&PublishDay=84&CID=" + str(args[1]) + "&IsDirectDelivery=&QualityType=&OrgStockStatus="
    req = requests.get(url)

    bs = BeautifulSoup(req.text, 'html.parser')

    book_html = bs.find_all("div", "ss_book_box")
    if not book_html:
        return "NULL"

    for book in book_html:
        temp = book.find("a", "bo_l")

        itemid = book.get('itemid')
        title = temp.find("b").string
        isbn = temp.get('href').split('=')[1].split('&')[0]
        stock = book.find("span", "ss_p4").find("b").string[1:]
        price = book.find("span", "ss_p2").find("b").string
        price = int(price.replace(",", "").replace("원", ""))

        temp_list.append(Book(title, isbn, itemid, stock, price, args[1]))

    return temp_list

# 정가 검색 및 ISBN 13 검색
def searchPrice(book):

    if sys.getrecursionlimit() != 100000:
        sys.setrecursionlimit(100000)

    try:
        url = "http://used.aladin.co.kr/shop/usedshop/wc2b_search.aspx?ActionType=1&SearchTarget=All&KeyWord=" + book.isbn

        req = requests.get(url)
        bs = BeautifulSoup(req.text, 'html.parser')

        temp = bs.find("table", {"id":"searchResult"})

        price = temp.find("td", "c2b_tablet3").string
        book.fixprice = int(price.replace(",", "").replace("원", ""))

        isbn_str = bs.find('table',id='searchResult').find('table').find('td').find_all('br')[-1].get_text()

        book.isbn13 = isbn_str.split(",")[0]
        if len(book.isbn13) != 13:
            book.isbn13 = ""

    except:
        book.fixprice = -1

    return book



'''
 ***********************
 **** 엑셀 출력 함수 ****
 ***********************
'''

def printExcel(offcode, book_len, workbook):
    # WORKSHEET CHECKING
    worksheet = workbook.get_worksheet_by_name(offcode_list[offcode])
    if not worksheet:
        worksheet = workbook.add_worksheet(offcode_list[offcode])

    header_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'fg_color': '#D9D9D9'})

    string_format = workbook.add_format({'align': 'center', 'font_size': 10, 'valign': 'vcenter'})

    # HEADER COLUMN
    worksheet.set_column(0, 0, 7)
    worksheet.set_column(1, 1, 50)
    worksheet.set_column(2, 3, 14)
    worksheet.set_column(4, 4, 8.5)
    worksheet.set_column(5, 5, 7)
    worksheet.set_column(6, 6, 15)
    worksheet.set_column(7, 7, 12)
    worksheet.set_column(8, 9, 8.5)
    worksheet.set_row(0, 19.5)

    # PRINT HEADER
    worksheet.write(0, 0, "번호", header_format)
    worksheet.write(0, 1, "책이름", header_format)
    worksheet.write(0, 2, "ISBN13", header_format)
    worksheet.write(0, 3, "ISBN10", header_format)
    worksheet.write(0, 4, "최저가", header_format)
    worksheet.write(0, 5, "재고", header_format)
    worksheet.write(0, 6, "카테고리", header_format)
    worksheet.write(0, 7, "Item ID", header_format)
    worksheet.write(0, 8, "정가", header_format)
    worksheet.write(0, 9, "My Price", header_format)

    # 데이터 출력
    idx = book_len
    for book in book_list:
        worksheet.set_row(idx, 18)
        worksheet.write(idx, 0, idx, string_format)
        worksheet.write(idx, 1, book.title, string_format)
        worksheet.write(idx, 2, book.isbn13, string_format)
        worksheet.write(idx, 3, book.isbn, string_format)
        worksheet.write(idx, 4, book.price, string_format)
        worksheet.write(idx, 5, book.stock, string_format)
        worksheet.write(idx, 6, cid_list[str(book.cid)], string_format)
        worksheet.write(idx, 7, book.itemid, string_format)
        worksheet.write(idx, 9, book.myprice, string_format)

        if book.fixprice == -1:
            worksheet.write(idx, 8, "절판", string_format)
        else:
            worksheet.write(idx, 8, book.fixprice, string_format)
        idx = idx+1



'''
 ************************
 ****** 제어용 함수 ******
 ************************
'''

# 실행용 함수 아님.
def crawl_run(offcode, cid, book_idx, factor_list, text_edit, pg_bar, process_cnt, workbook):

    if sys.getrecursionlimit() != 100000:
        sys.setrecursionlimit(100000)

    # 0. 변수 선언
    global book_list
    book_list = []

    print("[알림]", offcode_list[offcode], cid_list[cid], "도서 목록 크롤링 시작.")
    text_edit.append("[알림] " + offcode_list[offcode] + " " + cid_list[cid] + " 정보 수집 시작.")
    text_edit.repaint()

    # 1. 도서 목록 크롤링
    time1 = time.time()
    try:
        pool = Pool(processes=process_cnt)
        temp = pool.map(getBooks, getPages(offcode, cid))
        for a in temp:
            book_list.extend(a)
            pool.close()
            pool.join()
    except:
        text_edit.append("[실패] " + offcode_list[offcode] + " " + cid_list[cid] + " 페이지가 존재하지 않는 것 같습니다.")
        text_edit.repaint()
        return 0

    print("[알림] 도서 목록 크롤링 완료. (", round(time.time() - time1, 4), "sec )")
    text_edit.append("[알림] 도서 목록 크롤링 완료. : " + str(len(book_list)) + "권 ( " + str(round(time.time() - time1, 4)) + " sec )")
    print(len(book_list), "권")
    text_edit.repaint()
    pg_bar.setValue(pg_bar.value() + 1)
    pg_bar.repaint()

    # 2. 도서 정가 크롤링
    text_edit.append("[알림] 정가 크롤링 시작.")
    text_edit.repaint()
    time2 = time.time()
    pool = Pool(processes=process_cnt)
    book_list = pool.map(searchPrice, book_list)
    pool.close()
    pool.join()

    print("[알림] 정가 크롤링 완료. (", round(time.time() - time2, 4), "sec )")
    text_edit.append("[알림] 정가 크롤링 완료. ( " + str(round(time.time() - time2, 4)) + " sec )")
    text_edit.repaint()
    pg_bar.setValue(pg_bar.value() + 1)
    pg_bar.repaint()

    # 3. MyPrice 계산
    text_edit.append("[알림] MyPrice 계산 시작.")
    text_edit.repaint()
    time3 = time.time()
    for book in book_list:
        book.setMyPrice(factor_list)

    print("[알림] MyPrice 계산 완료. (", round(time.time() - time3, 4), "sec )")
    text_edit.append("[알림] MyPrice 계산 완료. ( " + str(round(time.time() - time3, 4)) + " sec )")
    text_edit.repaint()
    pg_bar.setValue(pg_bar.value() + 1)
    pg_bar.repaint()

    # 4. 엑셀에 출력
    text_edit.append("[알림] 엑셀 작성 시작.")
    text_edit.repaint()
    time4 = time.time()
    printExcel(offcode, book_idx + 1, workbook)

    print("[알림] 엑셀 작성 완료. (", round(time.time() - time4, 4), "sec )")
    text_edit.append("[알림] 엑셀 작성 완료. ( " + str(round(time.time() - time4, 4)) + " sec )")
    text_edit.repaint()
    pg_bar.setValue(pg_bar.value() + 1)
    pg_bar.repaint()


    print("[알림]", offcode_list[offcode], cid_list[cid], "수행 완료. ( 총", round(time.time() - time1, 4), "sec )")
    text_edit.append("[알림] " + offcode_list[offcode] + " " + cid_list[cid] + " 수행 완료. ( 총 " + str(round(time.time() - time1, 4)) + " sec )\n")
    text_edit.repaint()

    return len(book_list)

# 이 함수를 실행하세요.
def app_run(exe_offcode_list, exe_cid_list, factor_list, textedit, pgbar, processes):
    if sys.getrecursionlimit() != 100000:
        sys.setrecursionlimit(100000)

    pgvalue = len(exe_offcode_list) * len(exe_cid_list) * 4
    pgbar.setMaximum(pgvalue)

    i=1
    for offcode in exe_offcode_list:
        book_idx = 0
        workbook = xlsxwriter.Workbook(offcode_list[offcode] + ".xlsx")
        for cid in exe_cid_list:
            book_idx += crawl_run(offcode, cid, book_idx, factor_list, textedit, pgbar, processes, workbook)
        workbook.close()
        i+=1

    return True

'''
 **********************
 ****** GUI 관련 ******
 **********************
'''

class XDialog(QDialog, aladin_form.Ui_Dialog):
    def __init__(self):
        QDialog.__init__(self)
        self.setupUi(self)
        # 콤보 박스 아이템 추가
        self.comboBox.addItem("기본 (i3 이상 권장, 4)")
        self.comboBox.addItem("빠름 (i5 이상 권장, 8)")
        self.comboBox.addItem("빠름+ (i7 이상 권장, 16)")

        # 버튼 이벤트 핸들러
        self.btn_exit.clicked.connect(self.exitForm)
        self.btn_start.clicked.connect(self.startCrawl)

        # 프로그래스바 0 초기화
        self.progressBar.setValue(0)
        self.progressBar.repaint()

    # 종료 버튼
    def exitForm(self):
        QMessageBox.information(self, "종료", "프로그램을 종료합니다.")
        exit(1)

    # 시작 버튼
    def startCrawl(self):
        # 지점 체크 확인
        offcode_true_list = []
        if self.cb_sinsa.isChecked() == True:
            offcode_true_list.append('sinsa')
        if self.cb_gangnam.isChecked() == True:
            offcode_true_list.append('gangnam')
        if self.cb_geondae.isChecked() == True:
            offcode_true_list.append('geondae')
        if self.cb_nowon.isChecked() == True:
            offcode_true_list.append('nowon')
        if self.cb_daehakro.isChecked() == True:
            offcode_true_list.append('daehakro')
        if self.cb_suyu.isChecked() == True:
            offcode_true_list.append('suyu')
        if self.cb_sillim.isChecked() == True:
            offcode_true_list.append('sillim')
        if self.cb_sinchon.isChecked() == True:
            offcode_true_list.append('sinchon')
        if self.cb_yeonsinnae.isChecked() == True:
            offcode_true_list.append('yeonsinnae')
        if self.cb_jamsil.isChecked() == True:
            offcode_true_list.append('jamsil')
        if self.cb_sincheon.isChecked() == True:
            offcode_true_list.append('sincheon')
        if self.cb_jongno.isChecked() == True:
            offcode_true_list.append('jongno')
        if self.cb_hapjeong.isChecked() == True:
            offcode_true_list.append('hapjeong')
        if self.cb_dongtan.isChecked() == True:
            offcode_true_list.append('dongtan')
        if self.cb_bucheon.isChecked() == True:
            offcode_true_list.append('bucheon')
        if self.cb_buksuwon.isChecked() == True:
            offcode_true_list.append('buksuwon')
        if self.cb_bundang.isChecked() == True:
            offcode_true_list.append('bundang')
        if self.cb_yatap.isChecked() == True:
            offcode_true_list.append('yatap')
        if self.cb_sanbon.isChecked() == True:
            offcode_true_list.append('sanbon')
        if self.cb_suwon.isChecked() == True:
            offcode_true_list.append('suwon')
        if self.cb_ilsan.isChecked() == True:
            offcode_true_list.append('ilsan')
        if self.cb_hwajeong.isChecked() == True:
            offcode_true_list.append('hwajeong')
        if self.cb_sangmu.isChecked() == True:
            offcode_true_list.append('sangmu')
        if self.cb_gwangju.isChecked() == True:
            offcode_true_list.append('gwangju')
        if self.cb_daegu.isChecked() == True:
            offcode_true_list.append('daegu')
        if self.cb_sangin.isChecked() == True:
            offcode_true_list.append('sangin')
        if self.cb_daejeoncityhall.isChecked() == True:
            offcode_true_list.append('daejeoncityhall')
        if self.cb_daejeon.isChecked() == True:
            offcode_true_list.append('daejeon')
        if self.cb_killy.isChecked() == True:
            offcode_true_list.append('killy')
        if self.cb_deokcheon.isChecked() == True:
            offcode_true_list.append('deokcheon')
        if self.cb_dongbo.isChecked() == True:
            offcode_true_list.append('dongbo')
        if self.cb_centum.isChecked() == True:
            offcode_true_list.append('centum')
        if self.cb_ulsan.isChecked() == True:
            offcode_true_list.append('ulsan')
        if self.cb_gyesan.isChecked() == True:
            offcode_true_list.append('gyesan')
        if self.cb_jeonju.isChecked() == True:
            offcode_true_list.append('jeonju')
        if self.cb_cheonan.isChecked() == True:
            offcode_true_list.append('cheonan')
        if self.cb_cheongju.isChecked() == True:
            offcode_true_list.append('cheongju')
        if self.cb_guwol.isChecked() == True:
            offcode_true_list.append('guwol')

        # 카테고리 체크 확인
        cid_true_list = []
        if self.cb_1230.isChecked() == True:
            cid_true_list.append('1230')
        if self.cb_2551.isChecked() == True:
            cid_true_list.append('2551')
        if self.cb_1108.isChecked() == True:
            cid_true_list.append('1108')
        if self.cb_656.isChecked() == True:
            cid_true_list.append('656')
        if self.cb_50246.isChecked() == True:
            cid_true_list.append('50246')
        if self.cb_55890.isChecked() == True:
            cid_true_list.append('55890')
        if self.cb_4395.isChecked() == True:
            cid_true_list.append('4395')
        if self.cb_13789.isChecked() == True:
            cid_true_list.append('13789')
        if self.cb_336.isChecked() == True:
            cid_true_list.append('336')
        if self.cb_351.isChecked() == True:
            cid_true_list.append('351')
        if self.cb_170.isChecked() == True:
            cid_true_list.append('170')
        if self.cb_798.isChecked() == True:
            cid_true_list.append('798')
        if self.cb_1196.isChecked() == True:
            cid_true_list.append('1196')
        if self.cb_1237.isChecked() == True:
            cid_true_list.append('1237')
        if self.cb_2105.isChecked() == True:
            cid_true_list.append('2105')
        if self.cb_1.isChecked() == True:
            cid_true_list.append('1')
        if self.cb_74.isChecked() == True:
            cid_true_list.append('74')
        if self.cb_2030.isChecked() == True:
            cid_true_list.append('2030')
        if self.cb_987.isChecked() == True:
            cid_true_list.append('987')
        if self.cb_55889.isChecked() == True:
            cid_true_list.append('55889')
        if self.cb_517.isChecked() == True:
            cid_true_list.append('517')
        if self.cb_1137.isChecked() == True:
            cid_true_list.append('1137')
        if self.cb_8257.isChecked() == True:
            cid_true_list.append('8257')
        if self.cb_1383.isChecked() == True:
            cid_true_list.append('1383')
        if self.cb_1322.isChecked() == True:
            cid_true_list.append('1322')
        if self.cb_50245.isChecked() == True:
            cid_true_list.append('50245')

        # 멀티프로세싱
        if self.comboBox.currentIndex() == 1:
            processes = 8
        elif self.comboBox.currentIndex() == 2:
            processes = 16
        else:
            processes = 4

        try:
            # My Factor
            factor_list = []
            try:
                factor_list.append(float(self.factor1.text()))
                factor_list.append(int(self.factor2.text().replace("%", ""))/100)
                factor_list.append(int(self.factor3.text().replace("%", "")) / 100)
                factor_list.append(int(self.factor1_2.text()))
            except:
                QMessageBox.information(self, "에러", "데이터를 정확히 입력하십시오.")
                exit(1)

            # 실행
            start_time = time.time()
            self.textEdit.setText("[알림] 크롤링 시작\n")
            self.textEdit.append("[데이터] 가격변환율 : " + str(factor_list[0]) + " / 정가대비최대 : " + str(factor_list[1]) + "배 / 정가대비최소 : " + str(factor_list[2]) + "배 / 최저보장가 : " + str(factor_list[3]) + "원")
            self.textEdit.append("[데이터] 프로세스 개수 : " + str(processes) + "\n")
            self.textEdit.repaint()
            self.progressBar.setValue(0)

            if app_run(offcode_true_list, cid_true_list, factor_list, self.textEdit, self.progressBar, processes):
                self.progressBar.setValue(0)
                QMessageBox.information(self, "크롤링 완료", "소요시간 " + str(round(time.time() - start_time, 4)) + "sec")

        except Exception as emsg:
            QMessageBox.information(self, "에러", "에러 발생.\n에러 메세지 : " + str(emsg))
            exit(1)




'''
 ***********************
 ****** MAIN 함수 ******
 ***********************
'''

if __name__ == "__main__":
    if sys.platform.startswith('win'):
        multiprocessing.freeze_support()

    #global workbook

    #workbook = xlsxwriter.Workbook("aladin.xlsx")

    # GUI
    app = QApplication([])
    dlg = XDialog()
    dlg.show()
    app.exec_()

    sys.setrecursionlimit(100000)

    #workbook.close()