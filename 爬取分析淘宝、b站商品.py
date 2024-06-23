import tkinter as tk
import pandas
import openpyxl

t = tk.Tk()
t.geometry('1000x600')
t.title('商品数据分析')



def searching():
    if input_good.get()=='':
        label10=tk.Label(t,text='商品搜索不能为空！',font=('等线', 12),fg='red')
        label10.place(x=290,y=40)
        t.after(600,label10.destroy)
        return
    if(password.get()=='' or account.get()==''):
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='没有账号或密码，不能查询淘宝数据！')
        return
    taobao(input_good.get(),max_page.get())
    bilibili(input_good.get(),max_page.get())
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message='数据搜索完毕，可以进行数据分析')

def searching_taobao():
    if input_good.get()=='':
        label10=tk.Label(t,text='商品搜索不能为空！',font=('等线', 12),fg='red')
        label10.place(x=290,y=40)
        t.after(600,label10.destroy)
        return
    if(password.get()=='' or account.get()==''):
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='没有账号或密码，不能查询淘宝数据！')
        return
        
    taobao(input_good.get(),max_page.get())
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message='数据搜索完毕，可以进行数据分析')

def searching_bilibili():
    if input_good.get()=='':
        label10=tk.Label(t,text='商品搜索不能为空！',font=('等线', 12),fg='red')
        label10.place(x=290,y=40)
        t.after(600,label10.destroy)
        return
    bilibili(input_good.get(),max_page.get())
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message='数据搜索完毕，可以进行数据分析')

def taobao(good_name,max_page):
    from DrissionPage import ChromiumPage
    import re
    import json
    from DataRecorder import Recorder

    r = Recorder('淘宝搜索'+good_name+'的结果'+'.xlsx')
    r.add_data(['title','price','realSales','Shopname','procity'])
    r.record()

    #开始爬虫
    #1 打开浏览器
    page = ChromiumPage()

    #已经登录就不需要以下操作
    #2 进入淘宝登录页
    page.get('https://login.taobao.com/member/login.jhtml')
    #2.1 定位账号输入框
    page.ele('css:#fm-login-id').input(account.get())
    #2.2 定位密码输入框
    page.ele('css:#fm-login-password').input(password.get())
    #2.3 定位登录账号按钮
    page.ele('css:.fm-button.fm-submit.password-login').click()

    #3 输入我们要获取的商品名称 点击搜索
    page.get('https://www.taobao.com/')
    page.ele('css:#q').input(good_name)
    #3.2 点击搜索
    page.ele('css:.btn-search.tb-bg').click()
    #已经登录就不需要以下操作
    #page.ele('css:#fm-login-id').input("13343460922")
    #page.ele('css:#fm-login-password').input("hsqxfy911922")
    #page.ele('css:.fm-button.fm-submit.password-login').click()
    #3.1 找到输入框，输入商品名称，点击搜索

    #开启监听接口(网络面板的请求加载)
    page.listen.start('h5/mtop.relationrecommend.wirelessrecommend.recommend/2.0')
    global taobao_count
    taobao_count=0
    #4 开始等待 请求加载完毕 获取响应数据
    for p in range(1,max_page+1):
        response = page.listen.wait().response
        mtopjsonp = response.body
        #print(mtopjsonp)

        #5 正则表达式提取，将相应json数据 提取出来 转为字典
        mtopjson = re.findall(pattern='mtopjsonp\d+\((.*)\)',string=mtopjsonp)[0]
        #print(mtopjson)
        #得到实际的字典数据
        mtop_data = json.loads(mtopjson)

        #6. 取数据
        itemsArray = mtop_data['data']['itemsArray']
        for item in itemsArray:
            title = item['title']
            title = re.sub('<span class=H>','',title)
            title = re.sub('</span>','',title)
            price = item['price']
            realSales = item['realSales']
            procity = item['procity']
            Shopname = item['shopInfo']['title']
            #print(title,price,realSales,Shopname,procity)
            r.add_data([title,price,realSales,Shopname,procity])
            r.record()
            taobao_count+=1
        #下一页
        page.ele('css:.next-btn.next-small.next-btn-normal.next-pagination-item.next-next').click()

def bilibili(good_name,max_page):
    #pip install selenium==3.141.0
    #下载Chrome驱动(Chrome Drivers)   https://googlechromelabs.github.io/chrome-for-testing/#stable
    #将下载后压缩包中的exe文件添加到自己电脑上python的安装目录中

    from selenium import webdriver
    from time import sleep
    from DataRecorder import Recorder

    r = Recorder('会员购搜索'+good_name+'的结果'+'.xlsx')
    r.add_data(['title','price','wants'])

    #实例化浏览器对象
    driver = webdriver.Chrome()

    #访问网址
    driver.get('https://mall.bilibili.com/neul-next/index.html?page=flow_searchResult&goFrom=na&noTitleBar=1&fromType=0&from=mall_search_mall&category=&keyword='+good_name)
    driver.implicitly_wait(10)
    sleep(2)
    for i in range(1,max_page+1):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
    lists=driver.find_elements_by_xpath('//div/p[@class="name word-ellipsis-1 ab-test-title"]/span|//div/p[@class="name word-ellipsis-2"]|//div[@class="price"]|//span[@class="like"]')
    j=0
    global bilibili_count
    bilibili_count=0
    for i in lists:
        con = i.text
        if j%3==0:
            new=[]
            new.append(con)
        if j%3==1:
            new.append(con)
        if j%3==2:
            if '想要' in con:
                new.append(con)
                title=new[0]
                price=new[1]
                wants=new[2]
                r.add_data([title,price,wants])
                r.record()
            else:
                new.append('无')
                title=new[0]
                price=new[1]
                wants=new[2]
                r.add_data([title,price,wants])
                r.record()
                new=[]
                new.append(con)
                j=0
        j+=1
        bilibili_count+=1

def analysis():
    taobao_analysis(input_good.get(),min_price.get(),max_price.get(),amount.get())
    bilibili_analysis(input_good.get(),min_price.get(),max_price.get(),amount.get())
    import tkinter.messagebox
    if(input_good.get()==''):
        return
    tkinter.messagebox.showinfo(title='提示',message='数据分析完毕')

def analysis_taobao():
    taobao_analysis(input_good.get(),min_price.get(),max_price.get(),amount.get())
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message='数据分析完毕')

def analysi_bilibili():
    bilibili_analysis(input_good.get(),min_price.get(),max_price.get(),amount.get())
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message='数据分析完毕')

def taobao_graph():
    import matplotlib.pyplot as plt
    # 设置中文显示
    if 'taobao_df_sorted_ascending' not in globals():
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='没有分析数据，请先进行数据分析')
        return
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False


    plt.figure(figsize=(24,12))
    plt.barh(taobao_name,taobao_price,label='价格',linewidth=1)
    plt.xlabel('prices')
    plt.ylabel('name')
    plt.title('淘宝搜索'+input_good.get()+'价格分析情况')
    plt.tick_params(axis='y', labelsize=9)
    plt.show()

def taobao_output():
    if 'taobao_df_sorted_ascending' not in globals():
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='没有分析数据，请先进行数据分析')
        return
    taobao_df_sorted_ascending.to_excel('经筛选后的淘宝搜索'+input_good.get()+'价格分析情况.xlsx',index=False)
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='导出成功',message='文件已成功导出')

def taobao_analysis(good_name,min_price,max_price,amount):
    if input_good.get()=='':
        label10=tk.Label(t,text='商品搜索不能为空！',font=('等线', 12),fg='red')
        label10.place(x=290,y=40)
        t.after(600,label10.destroy)
        return
    import pandas as pd
    import numpy as np


    search_content=good_name
    try:
        cur_file = pd.read_excel('淘宝搜索'+search_content+'的结果.xlsx')
        
    except:
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='找不到淘宝搜索'+search_content+'的结果.xlsx文件')
        return

    taobao_result.set('')
    data = np.array(cur_file) #转换为numpy数组
    global taobao_name
    taobao_name=[]
    global taobao_price
    taobao_price=[]
    global taobao_realSale
    taobao_realSale=[]
    global taobao_Shopname
    taobao_Shopname=[]
    global taobao_procity
    taobao_procity=[]

    count_max=amount
    count=0
    del_count=0
    right_count=0

    for ls in data:
        if ls[1]>=min_price and ls[1]<=max_price:
            taobao_name.append(ls[0])
            taobao_price.append(ls[1])
            taobao_realSale.append(ls[2])
            taobao_Shopname.append(ls[3])
            taobao_procity.append(ls[4])
            right_count+=1
            lb1.insert(tk.END, ls[0])
            lb1.insert(tk.END, ls[1])
            lb1.insert(tk.END, ls[2])
            lb1.insert(tk.END, ls[3])
            lb1.insert(tk.END, ls[4])
            
        else:
            del_count+=1
        count+=1
        if count==count_max:
            break

    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message="淘宝已剔除%d个数据"%(del_count))

    obj={'shop':taobao_name,'price':taobao_price,'realSale':taobao_realSale,'Shopname':taobao_Shopname,'procity':taobao_procity}
    global taobao_df_obj
    taobao_df_obj=pd.DataFrame(obj)
    taobao_min=taobao_df_obj['price'].min()
    taobao_max=taobao_df_obj['price'].max()
    taobao_mean=taobao_df_obj['price'].mean()
    label12.config(text='最高价格为:'+str(taobao_max))
    label13.config(text='最低价格为:'+str(taobao_min))
    label14.config(text='平均价格为:'+str(taobao_mean))
    global taobao_df_sorted_ascending
    taobao_df_sorted_ascending = taobao_df_obj.sort_values(by='price', ascending=True)

def bilibili_graph():
    if 'bilibili_df_sorted_ascending' not in globals():
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='没有分析数据，请先进行数据分析')
        return
    import matplotlib.pyplot as plt
        # 设置中文显示
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False


    plt.figure(figsize=(24,12))
    plt.barh(bilibili_name,bilibili_price,label='价格',linewidth=1)
    plt.xlabel('prices')
    plt.ylabel('titles')
    plt.title('哔哩哔哩会员购搜索'+input_good.get()+'分析情况')
    plt.tick_params(axis='y', labelsize=9)
    plt.show()

def bilibili_output():
    if 'bilibili_df_sorted_ascending' not in globals():
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='没有分析数据，请先进行数据分析')
        return
    bilibili_df_sorted_ascending.to_excel('经筛选后的会员购搜索'+input_good.get()+'价格分析情况.xlsx',index=False)
    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='导出成功',message='文件已成功导出')

def bilibili_analysis(good_name,min_price,max_price,amount):
    import pandas as pd
    import numpy as np
    import matplotlib.pyplot as plt
    import re

    if input_good.get()=='':
        label10=tk.Label(t,text='商品搜索不能为空！',font=('等线', 12),fg='red')
        label10.place(x=290,y=40)
        t.after(600,label10.destroy)
        return

    search_content=good_name
    
    def extract_numbers(s):
        return [int(digit) for digit in re.findall(r'\d+', s)]

    try:
        cur_file = pd.read_excel('会员购搜索'+good_name+'的结果.xlsx')
        
    except:
        import tkinter.messagebox
        tkinter.messagebox.showerror(title='错误',message='找不到哔哩哔哩搜索'+search_content+'的结果.xlsx文件')
        return

    bilibili_result.set('')

    data = np.array(cur_file) #转换为numpy数组
    global  bilibili_name
    bilibili_name=[]
    global  bilibili_price
    bilibili_price=[]
    global  bilibili_wants
    bilibili_wants=[]

    count_max=amount
    count=0
    del_count=0
    right_count=0
    price_min=min_price
    price_max=max_price

    for ls in data:
        try:
            numbers = extract_numbers(ls[1])
            correct_price=numbers[0]
        except:
            break
        if correct_price>=price_min and correct_price<=price_max:
            bilibili_name.append(ls[0])
            bilibili_price.append(correct_price)
            bilibili_wants.append(ls[2])
            right_count+=1
            lb2.insert(tk.END, ls[0])
            lb2.insert(tk.END, ls[1])
            lb2.insert(tk.END, ls[2])
        else:
            del_count+=1
        count+=1
        if count==count_max:
            break

    import tkinter.messagebox
    tkinter.messagebox.showinfo(title='提示',message="会员购已剔除%d个数据"%(del_count))


    obj={'shop':bilibili_name,'price':bilibili_price,'wants':bilibili_wants}

    global bilibili_df_obj
    bilibili_df_obj=pd.DataFrame(obj)
    bilibili_min=bilibili_df_obj['price'].min()
    bilibili_max=bilibili_df_obj['price'].max()
    bilibili_mean=bilibili_df_obj['price'].mean()
    label15.config(text='最高价格为:'+str(bilibili_max))
    label16.config(text='最低价格为:'+str(bilibili_min))
    label17.config(text='平均价格为:'+str(bilibili_mean))
    global bilibili_df_sorted_ascending
    bilibili_df_sorted_ascending = bilibili_df_obj.sort_values(by='price', ascending=True)




input_good = tk.StringVar()
input_good.set('')
max_page = tk.IntVar()  
max_page.set(3)
max_price= tk.IntVar()
max_price.set(99999)
min_price= tk.IntVar()
min_price.set(0)
taobao_result= tk.StringVar()
taobao_result.set('')
bilibili_result=tk.StringVar()
bilibili_result.set('')
amount= tk.IntVar()
amount.set(50)
account = tk.StringVar()
account.set('')
password = tk.StringVar()
password.set('')

label1=tk.Label(t,text='输入要搜索的商品名称：', font=('等线', 16, 'bold'))
label1.grid(row=0,column=0,padx=5,pady=5,sticky='e')
ent1 = tk.Entry(t,width=30,textvariable=input_good)
ent1.grid(row=0,column=1,padx=5,pady=5,sticky="e")

button1=tk.Button(t,text='搜索数据',command=searching,font=('等线', 12, 'bold'),height=2,width=10)
button1.place(x=425,y=60)
button5=tk.Button(t,text='仅搜索淘宝',command=searching_taobao,font=('等线', 12),height=1,width=10)
button5.place(x=535,y=69)
button6=tk.Button(t,text='仅搜索会员购',command=searching_bilibili,font=('等线', 12),height=1,width=10)
button6.place(x=645,y=69)

label3=tk.Label(t,text='设置搜索数据的页数（数据量）：',font=('等线', 12))
label3.grid(row=0,column=2,padx=5,pady=5,sticky='e')
scale = tk.Scale(variable = max_page,from_ = 1, to = 50, resolution=1,orient = 'horizontal',length=250)  
scale.grid(row=0,column=3,padx=5,pady=5,sticky='e')

button2=tk.Button(t,text='分析数据',command=analysis,font=('等线', 12, 'bold'),height=2,width=10)
button2.place(x=425,y=120)
button7=tk.Button(t,text='仅分析淘宝',command=analysis_taobao,font=('等线', 12),height=1,width=10)
button7.place(x=535,y=120)
button8=tk.Button(t,text='仅分析会员购',command=analysi_bilibili,font=('等线', 12),height=1,width=10)
button8.place(x=645,y=120)


button3=tk.Button(t,text='导出淘宝筛选数据为表格',command=taobao_output,font=('等线', 12, 'bold'),height=1,width=20)
button3.place(x=50,y=180)
button4=tk.Button(t,text='显示淘宝筛选数据柱状图',command=taobao_graph,font=('等线', 12, 'bold'),height=1,width=20)
button4.place(x=250,y=180)

button3=tk.Button(t,text='导出b站会员购筛选数据为表格',command=bilibili_output,font=('等线', 12, 'bold'),height=1,width=24)
button3.place(x=480,y=180)
button4=tk.Button(t,text='显示b站会员购筛选数据柱状图',command=bilibili_graph,font=('等线', 12, 'bold'),height=1,width=24)
button4.place(x=720,y=180)

label8=tk.Label(t,text='设置最低金额：', font=('等线', 12,))
label8.place(x=760,y=82)
ent2 = tk.Entry(t,width=10,textvariable=min_price)
ent2.place(x=870,y=82)
label9=tk.Label(t,text='设置最高金额：', font=('等线', 12,))
label9.place(x=760,y=112)
ent3 = tk.Entry(t,width=10,textvariable=max_price)
ent3.place(x=870,y=112)

label11=tk.Label(t,text='设置分析数据的量', font=('等线', 12,))
label11.place(x=155,y=105)
scale2 = tk.Scale(variable = amount,from_ = 5, to = 400, resolution=5,orient = 'horizontal',length=300)  
scale2.place(x=70,y=127)

label6=tk.Label(t,text='淘宝筛选的数据：', font=('等线', 16, 'bold'))
label6.place(x=40,y=225)
global lb1
lb1=tk.Listbox(t,listvariable=taobao_result,width=60,height=15)
lb1.place(x=40,y=265)
global label12
global label13
global label14
label12=tk.Label(t,text='', font=('等线', 12,))
label12.place(x=75,y=545)
label13=tk.Label(t,text='', font=('等线', 12,))
label13.place(x=255,y=545)
label14=tk.Label(t,text='', font=('等线', 12,))
label14.place(x=125,y=565)

label7=tk.Label(t,text='哔哩哔哩会员购筛选的数据：', font=('等线', 16, 'bold'))
label7.place(x=500,y=225)
global lb2
lb2=tk.Listbox(t,listvariable=bilibili_result,width=60,height=15)
lb2.place(x=500,y=265)
label15=tk.Label(t,text='', font=('等线', 12,))
label15.place(x=575,y=545)
label16=tk.Label(t,text='', font=('等线', 12,))
label16.place(x=755,y=545)
label17=tk.Label(t,text='', font=('等线', 12,))
label17.place(x=635,y=565)

label18=tk.Label(t,text='输入你的淘宝账号：', font=('等线', 12,))
label18.place(x=5,y=55)
ent4 = tk.Entry(t,width=25,textvariable=account)
ent4.place(x=155,y=55)
label19=tk.Label(t,text='输入你的淘宝密码：', font=('等线', 12,))
label19.place(x=5,y=80)
ent5 = tk.Entry(t,width=25,textvariable=password)
ent5.place(x=155,y=80)


t.mainloop()

