from tkinter import ttk, scrolledtext, filedialog, messagebox
from tkinter import Tk, Frame, LabelFrame, Button, Entry, Scale, Canvas, StringVar
from tkinter import END, HORIZONTAL, CENTER
import base64
from icon import Icon
import os
from datetime import datetime
import time
import requests
import threading
import xlrd
from queue import Queue


class Download:

    def __init__(self, root):

        self.plyNoLst = []
        self.download_path = os.path.join(os.path.expanduser('~'), 'Desktop', '电子保单-{}'.format(datetime.now().strftime('%Y%m%d'))).replace('\\', '/')
        root.title('电子保单下载')
        # 读取 icon.py 中的图标文件
        with open('tmp.ico', 'wb') as tmp:
            tmp.write(base64.b64decode(Icon().img))
        root.iconbitmap('tmp.ico')
        os.remove('tmp.ico')
        screenWidth = root.winfo_screenwidth()  # 获取显示区域的宽度
        screenHeight = root.winfo_screenheight()  # 获取显示区域的高度
        width, height = 381, 592
        left = int((screenWidth - width) / 2)
        top = int((screenHeight - height) / 2) - 50
        root.geometry('{width}x{height}+{left}+{top}'.format(width=width, height=height, left=left, top=top))
        root.resizable(0, 0)  # 设置窗口宽高固定
        self.input_plyNo(root)
        self.set_download_path(root)
        self.download_config(root)
        self.log_frame(root)


    ## 保单号导入
    def input_plyNo(self, root):
        input_plyNo_frame = LabelFrame(root, text='保单号导入')
        self.ctrl_v_input = Entry(input_plyNo_frame, width=40, justify=CENTER)
        self.ctrl_v_input.grid(row=0, column=0, padx=3, columnspan=1)
        Button(input_plyNo_frame, text='添加', width=4, padx=0, pady=0, command=self.add1plyNo).grid(row=0, column=1, columnspan=1, padx=3)
        Button(input_plyNo_frame, text='删除', width=4, padx=0, pady=0, command=self.remove1plyNo).grid(row=0, column=2, columnspan=1, padx=3)
        self.var_txt_path = StringVar()
        txt_input = Entry(input_plyNo_frame, textvariable=self.var_txt_path, width=40, justify=CENTER, state='readonly')
        txt_input.grid(row=1, column=0, padx=3, columnspan=1)
        Button(input_plyNo_frame, text='.txt导入', width=10, padx=0, pady=0, command=self.txt_input).grid(row=1, column=1, columnspan=2, padx=3, pady=3)
        self.var_excel_path = StringVar()
        excel_input = Entry(input_plyNo_frame, textvariable=self.var_excel_path, width=40, justify=CENTER, state='readonly')
        excel_input.grid(row=2, column=0, padx=3, columnspan=1)
        Button(input_plyNo_frame, text='Excel导入', width=10, padx=0, pady=0, command=self.excel_input).grid(row=2, column=1, columnspan=2, padx=3)
        self.var_plyNoLst_info = StringVar()
        self.show_plyNo = Entry(input_plyNo_frame, textvariable=self.var_plyNoLst_info, width=40, justify=CENTER, state='readonly')
        self.show_plyNo.grid(row=3, column=0, padx=3, columnspan=1)
        self.print_plyNoLst()
        Button(input_plyNo_frame, text='查看', width=4, padx=0, pady=0, command=self.show_all_plyNo).grid(row=3, column=1, columnspan=1, padx=3, pady=3)
        Button(input_plyNo_frame, text='清空', width=4, padx=0, pady=0, command=self.remove_all_plyNo).grid(row=3, column=2, columnspan=1, padx=3, pady=3)

        input_plyNo_frame.place(x=2, y=2)


    def txt_input(self):
        try:
            filename = filedialog.askopenfilename(title='打开文件', filetypes=[('文本文档', '*.txt')])
            self.var_txt_path.set(filename)
            if filename != '':
                try:
                    with open(filename, 'r', encoding='utf-8') as f:
                        temp = f.readlines()
                except:
                    with open(filename, 'r', encoding='gbk') as f:
                        temp = f.readlines()
                for i in range(len(temp)):
                    plyNo = temp[i].strip()
                    if plyNo not in self.plyNoLst and len(plyNo) > 10:
                        self.plyNoLst.append(plyNo)
                self.print_plyNoLst()
        except Exception as e:
            self.log_info_window['state'] = 'normal'
            self.log_info_window.insert(END, '{}\n'.format(e))
            self.log_info_window['state'] = 'disabled'
            self.log_info_window.see(END)


    def excel_input(self):
        try:
            filename = filedialog.askopenfilename(title='打开文件', filetypes=[('Excel 工作簿', '*.xls *.xlsx')])
            self.var_excel_path.set(filename)
            if filename != '':
                workbook = xlrd.open_workbook(filename)
                worksheet = workbook.sheet_by_index(0)
                need_column = '保单号'
                columns = worksheet.row_values(0)
                col_data = [x.strip() for x in list(set(worksheet.col_values(columns.index(need_column))))]
                for i in range(len(col_data)):
                    plyNo = col_data[i]
                    if plyNo not in self.plyNoLst and len(plyNo) > 10:
                        self.plyNoLst.append(plyNo)
                self.print_plyNoLst()
        except Exception as e:
            self.log_info_window['state'] = 'normal'
            self.log_info_window.insert(END, '{}\n'.format(e))
            self.log_info_window['state'] = 'disabled'
            self.log_info_window.see(END)


    def print_plyNoLst(self):
        if len(self.plyNoLst) == 0:
            self.var_plyNoLst_info.set('请导入待下载保单号')
        else:
            self.var_plyNoLst_info.set('{} 个电子保单待下载'.format(len(self.plyNoLst)))


    def add1plyNo(self):
        plyNo = self.ctrl_v_input.get().strip()
        if plyNo not in self.plyNoLst and len(plyNo) > 10:
            self.plyNoLst.append(plyNo)
        self.print_plyNoLst()


    def remove1plyNo(self):
        plyNo = self.ctrl_v_input.get().strip()
        if plyNo in self.plyNoLst:
            self.plyNoLst.remove(plyNo)
        self.print_plyNoLst()


    def show_all_plyNo(self):
        self.log_info_window['state'] = 'normal'
        self.log_info_window.insert(END, '\n---------待下载保单号---------\n')
        self.log_info_window['state'] = 'disabled'
        self.log_info_window.see(END)
        if not len(self.plyNoLst):
            self.log_info_window['state'] = 'normal'
            self.log_info_window.insert(END, '            （空）\n')
            self.log_info_window['state'] = 'disabled'
        for i in range(len(self.plyNoLst)):
            self.log_info_window['state'] = 'normal'
            self.log_info_window.insert(END, self.plyNoLst[i]+'\n')
            self.log_info_window['state'] = 'disabled'
            self.log_info_window.see(END)
        self.log_info_window['state'] = 'normal'
        self.log_info_window.insert(END, '------------------------------\n')
        self.log_info_window['state'] = 'disabled'
        self.log_info_window.see(END)
        

    def remove_all_plyNo(self):
        self.plyNoLst = []
        self.print_plyNoLst()


    ## 下载目录
    def set_download_path(self, root):
        download_path_frame = LabelFrame(root, text='下载目录')
        self.var_download_path = StringVar()
        self.var_download_path.set(self.download_path)
        download_path_info = Entry(download_path_frame, textvariable=self.var_download_path, width=40)
        download_path_info.grid(row=0, column=0, padx=3, columnspan=1)
        Button(download_path_frame, text='更改', width=4, padx=0, pady=0, command=self.change_download_path).grid(row=0, column=1, columnspan=1, padx=3, pady=3)
        Button(download_path_frame, text='打开', width=4, padx=0, pady=0, command=self.open_download_path).grid(row=0, column=2, columnspan=1, padx=3, pady=3)
        download_path_frame.place(x=2, y=150)


    def change_download_path(self):
        self.var_download_path.set(filedialog.askdirectory())

    def open_download_path(self):
        self.download_path = self.var_download_path.get()
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        os.startfile(self.download_path)


    ## 下载线程
    def download_config(self, root):
        download_config_frame = LabelFrame(root, text='下载线程数')
        self.var_thread_num = Scale(download_config_frame, showvalue=True, from_=1, to=50, length=283, orient=HORIZONTAL)
        self.var_thread_num.grid(row=0, column=0, padx=1, pady=4, columnspan=1)
        Canvas(download_config_frame, width=77, height=0).grid(row=0, column=1, padx=1, pady=1, columnspan=2)
        Button(root, text='开始下载', width=10, padx=0, pady=0, command=self.start2download).place(x=297, y=248)
        download_config_frame.place(x=2, y=208)


    def start2download(self):
        self.download_erro_lst = []
        self.download_finish_count = 0
        self.data_num = len(self.plyNoLst)
        self.startTime_G = time.time()  # Global
        self.plyNoQueue = Queue(self.data_num)
        for i in range(self.data_num):
            self.plyNoQueue.put(self.plyNoLst[i])
        thread_num = int(self.var_thread_num.get())
        self.download_path = self.var_download_path.get()
        if not os.path.exists(self.download_path):
            os.makedirs(self.download_path)
        self.mutex_lock = threading.Lock()
        for i in range(thread_num):  # 线程数
            t = threading.Thread(target=self.download_one_time, name='LoopThread' + str(i))
            t.setDaemon(True)
            t.start()


    ## 日志窗口
    def log_frame(self, root):
        log_frame = LabelFrame(root, text='日志')
        self.log_info_window = scrolledtext.ScrolledText(log_frame, width=49, height=19, state='disabled')
        self.log_info_window.grid(row=0, column=0, padx=5, columnspan=500)
        Button(log_frame, text='导出', width=4, padx=0, pady=0, relief='ridge', command=self.output_log).grid(row=1, column=497, pady=3, columnspan=1)
        Button(log_frame, text='清空', width=4, padx=0, pady=0, relief='ridge', command=self.clean_log).grid(row=1, column=498, pady=3, columnspan=1)
        Button(log_frame, text='关于', width=4, padx=0, pady=0, relief='ridge', command=self.about).grid(row=1, column=499, pady=3, columnspan=1)
        log_frame.place(x=2, y=284)


    def output_log(self):
        log_info = self.log_info_window.get(1.0, END)
        log_save_name = filedialog.asksaveasfilename(title='保存日志', initialfile='log.txt')
        if log_save_name != '':
            with open(log_save_name, 'w', encoding='utf-8') as f:
                f.write(log_info)
            self.log_info_window['state'] = 'normal'
            self.log_info_window.insert(END, '\n{} 日志导出成功！\n\n'.format(datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            self.log_info_window['state'] = 'disabled'
            self.log_info_window.see(END)

    def clean_log(self):
        self.log_info_window['state'] = 'normal'
        self.log_info_window.delete(1.0, END)
        self.log_info_window['state'] = 'disabled'


    def about(self):
        about_info = '使用方法：\n1 .txt 文件导入：每行一个保单号\n2 Excel 文件导入：第一个sheet中有一列“保单号”且没有重名列\n\nE-POLICY DOWNLOAD FOR AIC\nmrmmmt_ 制作 v1.1 (Build 20210416)\nE-mail：mmmt123321@126.com'
        messagebox.showinfo('关于', about_info)


    ## 保单下载
    def download_by_policyId(self, plyNo):
        url = 'http://www.alltrust.com.cn/person/download/policyId/{}'.format(plyNo)

        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9','Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Host': 'www.alltrust.com.cn',
            'Referer': 'http://www.alltrust.com.cn/new/policyAndClaim/query/type/policy/index',
            'Upgrade-Insecure-Requests': '1', 
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
        }

        save_name = self.download_path + '/' + plyNo + '.pdf'
        start_time = time.time()
        while True:
            if time.time() - start_time > 30:
                self.download_erro_lst.append(plyNo + '.pdf' + '  下载超时')
                os.remove(save_name)
                flag = '失败'
                break
            r = requests.get(url, headers=headers)
            with open(save_name, 'wb') as f:
                f.write(r.content)
            if os.path.getsize(save_name) > 150000:
                flag = '成功'
                break
        
        self.mutex_lock.acquire()
        self.download_finish_count += 1
        self.log_info_window['state'] = 'normal'
        self.log_info_window.insert(END, '{0}/{1} {2} {3}.pdf {4}\n'.format(self.download_finish_count, self.data_num, flag, plyNo, self.remainingTime()))
        self.log_info_window['state'] = 'disabled'
        self.log_info_window.see(END)
        self.mutex_lock.release()
        if self.download_finish_count == self.data_num:
            self.log_info_window['state'] = 'normal'
            self.log_info_window.insert(END, '\n下载完成，共下载{}单，{}单下载超时，共耗时{:.2f}秒'.format(self.download_finish_count, len(self.download_erro_lst), time.time()-self.startTime_G)+'\n')
            self.log_info_window['state'] = 'disabled'
            self.log_info_window.see(END)
            if len(self.download_erro_lst) > 0:
                for i, download_erro_info in enumerate(self.download_erro_lst):
                    self.log_info_window['state'] = 'normal'
                    self.log_info_window.insert(END, str(i+1)+'  '+download_erro_info+'\n')
                    self.log_info_window['state'] = 'disabled'
                    self.log_info_window.see(END)


    def download_one_time(self):
        while (not self.plyNoQueue.empty()):
            plyNo = self.plyNoQueue.get()
            self.download_by_policyId(plyNo)


    def remainingTime(self):
        '''计算剩余时间'''
        timeSpent = time.time() - self.startTime_G
        timeRemaining = int(timeSpent * (self.data_num-self.download_finish_count) / self.download_finish_count)
        s = timeRemaining % 60
        m = (timeRemaining - s) % 3600 // 60
        h = (timeRemaining - s - m*60) // 3600
        return "{h}:{m}:{s}".format(h=h, m=str(m).zfill(2), s=str(s).zfill(2))


if __name__ == '__main__':
    root = Tk()
    Download(root)
    root.mainloop()

