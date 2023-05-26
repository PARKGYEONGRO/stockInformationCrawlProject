#Need install pandas, lxml, selenium
import os
from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
import pandas as pd
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By


folderSelected=() # 전역변수


# 프로그램 메인 창
main = Tk()
main.title("PGR_StockInformationCrowl")
main.resizable(False, False) # 창 크기 고정
main.geometry("1020x700+100+100")
main['bg'] = '#313338'


# Cmd explain
def explain():
    msgbox.showinfo('StockInformationCrowl', '1. 조회할 항목 입력\n2. 저장 경로 설정\n3. 시작 버튼 클릭\n파일 열기로 간단하게 저장된 Excel 파일 확인 가능\n* 크롤링 사이트는 네이버 시가총액 입니다.*')

# Cmd start
def start():
    global folderSelected
    if len(folderSelected) == 0:
        msgbox.showwarning("경고", "저장 경로를 선택하세요")
        labelFile2['fg'] = '#fb7e7e'
        return

    
    # 1. 페이지 이동
    browser = webdriver.Chrome()
    browser.maximize_window()
    url ='https://finance.naver.com/sise/sise_market_sum.nhn?&page='
    browser.get(url)


    # 2. 조회 항목 초기화 (체크 해제)
    checkboxes = browser.find_elements(By.NAME, 'fieldIds')
    for checkbox in checkboxes:
        if checkbox.is_selected(): # 체킹 상황
            checkbox.click() # 클릭 (체크 해제)

    # 3. 조회 항목 설정 (원하는 항목 체킹)
    for checkbox in checkboxes:
        parent = checkbox.find_element(By.XPATH, '..')  # 부모 element
        label = parent.find_element(By.TAG_NAME, 'label')
        if label.text in queryItem: # label.text에 원하는 항목이 있을경우
            checkbox.click() # 체킹
    # 4. 적용하기
    btn_apply = browser.find_element(By.XPATH, '//a[@href="javascript:fieldSubmit()"]')
    btn_apply.click()

    for idx in range(1,40): # 1이상 40 미만 페이지 반복
        # 사전 작업 : 페이지 이동
        browser.get(url + str(idx)) # &page=1~40   


        # 5. 데이터 추출
        df = pd.read_html(browser.page_source)[1]
        df.dropna(axis='index', how='all', inplace=True)
        df.dropna(axis='columns', how='all', inplace=True)


        # 더 이상 가져올 데이터가 없을 때
        if len(df) == 0:
            break


        # 6. 파일 저장
        fileName = 'stock.csv'
        destPath = folderSelected + f'\{fileName}'
        if os.path.exists(destPath): # 파일 있다면 헤더 제외
            df.to_csv(destPath, encoding='utf-8-sig', index=False, mode='a', header=False)
        else: # 헤더 포함
            df.to_csv(destPath, encoding='utf-8-sig', index=False)

        print(f'{idx} 페이지 완료')


    browser.quit() # 브라우저 종료

# Cmd close
def close():
    closeAnswer = msgbox.askquestion('StockInformationCrowl', '프로그램을 종료하시겠습니까?')
    if closeAnswer == 'yes':
        main.destroy()

def fileDialog():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="가져오기",
                                          filetype=(('CSV files', '*.csv'),("xlsx files", "*.xlsx"),("All Files", "*.*")))
    labelFile1["text"] = filename
    labelFile1['fg'] = '#6aa2ff'
    filePath = labelFile1["text"]
    try:
        excel_filename = r"{}".format(filePath)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        msgbox.showerror("오류", "선택한 파일이 잘못되었습니다.")
        return None
    except FileNotFoundError:
        msgbox.showerror("오류", f"다음과 같은 파일이 없습니다. {filePath}")
        labelFile1["text"] = 'No File Selected'
        labelFile1['fg'] = '#fb7e7e'
        return None

    clearData()

    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def clearData():
    tv1.delete(*tv1.get_children())
    return None

def browserStoragePath():
    global folderSelected
    folderSelected = filedialog.askdirectory()
    if folderSelected == '': # 사용자가 취소를 눌를 때
        labelFile2['fg'] = '#fb7e7e'
        return
    else:
        labelFile2['fg'] ='#6aa2ff'
    labelFile2["text"] = folderSelected

def insertQueryItem():
    global queryItem
    queryItem =[]
    queryItem = insertArea.get().split()
    if len(queryItem) >= 1:
        settingStatus['text'] = '세팅 O'
        settingStatus['fg'] = '#6aa2ff'
    else:
        settingStatus['text'] = '세팅 X'
        settingStatus['fg'] = '#fb7e7e'
    insertArea.delete(0, END)
    print(queryItem) # 터미널로 입력된 리스트 확인

# Enter키를 누르면 insertQueryItem()함수를 실행시키는 함수를 따로 정의, Enter키는 event이기 때문
def Enterstart(event):
    insertQueryItem()


mainFrame = Frame(main,height=500, width=1000, relief='solid')
mainFrame.place(x=10,y=10)
frame1 = ttk.LabelFrame(mainFrame, text="Excel Data")
frame1.place(height=500, width=1000)
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # Full 사이즈
treescrolly = Scrollbar(frame1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx = Scrollbar(frame1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


# 조회 항목 세팅 프레임
queryItemSetingFrame = LabelFrame(main, text='조회 항목 세팅', bg='#313338', fg='white')
queryItemSetingFrame.place(height=90, x=10, y=510, width=370)
labelQISF = Label(queryItemSetingFrame, text='띄어쓰기로 구분\n(예시 : 매출액 영업이익 자산총계)', bg='#313338', fg='white', justify='left')
labelQISF.place(rely=0, relx=0.01)
insertArea = Entry(queryItemSetingFrame, width=35)
insertArea.place(rely=0.525, relx=0.01)
settingStatus = Label(queryItemSetingFrame, text="세팅 전", bg='#313338', fg="yellow")
settingStatus.place(rely=0.5, relx=0.7)
insertBtn= Button(queryItemSetingFrame, text="입력", relief='flat', bg='#3d3e42', fg='white', width=6 ,command=insertQueryItem)
insertBtn.place(rely=0.45, relx=0.835) 
insertArea.bind('<Return>',Enterstart) # 키패드 Enter 키 바인드


# File 프레임
fileFrame = LabelFrame(main, text='File', bg='#313338', fg='white')
fileFrame.place(height=80, x=10, y=600, width=425)

# 파일 열기, 저장경로
labelFile1 = Label(fileFrame, text='No File Selected', bg='#313338', fg='white')
labelFile1.place(rely=0, relx=0.01)
labelFile2 = Label(fileFrame, text='No StoragePath', bg='#313338', fg='white')
labelFile2.place(rely=0.5, relx=0.01)
button1 = Button(fileFrame, text='파일 열기', relief='flat', bg='#3d3e42', fg='white', command=lambda: fileDialog())
button1.place(rely=0, relx=0.825)
button2 = Button(fileFrame, text='저장 경로', relief='flat', bg='#3d3e42', fg='white', command=lambda: browserStoragePath())
button2.place(rely=0.5, relx=0.825)


# 하위 프레임
footerFrame = Frame(main, bg='#313338')
footerFrame.place(x=820,y=630)

# 하위 버튼
closeBtn = Button(footerFrame, text="끝내기", relief='flat', bg='#3d3e42', fg='white', command=close)
closeBtn.grid(row=0,column=3,ipadx=10,ipady=10)
startBtn = Button(footerFrame, text="시작", relief='flat', bg='#3d3e42', fg='white', command=start)
startBtn.grid(row=0,column=2,ipadx=10,ipady=10,padx=5,pady=5)
explainBtn = Button(footerFrame, text="설명", relief='flat', bg='#3d3e42', fg='white', command=explain)
explainBtn.grid(row=0,column=1,ipadx=10,ipady=10)


main.mainloop()