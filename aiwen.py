# -*- coding:utf-8 -*-
import urllib
import urllib2
import re
import thread
import time
import xlsxwriter
import sys
import threading
import random

keywords=u"旅游" #设置搜索关键词
start_num=1 #设置需要爬取的起始页面
end_num=5 #设置需要爬取的结束页面

split_list=['\n',u'、',u'。',u'！',u'？',u'----','?','!']
str_list=[u'★',u'——',u'~',u'￥',u'☆','>',u'#',u'=',u'【',u'】',u'《',u'》',u'；', u'O(∩_∩)O',u'( ⊙ o ⊙ )',u'*',u'链接']
'''
header={'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) \
        AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
'''
headers_list=['Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',\
              'Opera/9.80 (Windows NT 5.1; U; zh-cn) Presto/2.9.168 Version/11.50',\
              'Mozilla/5.0 (Windows NT 5.1; rv:5.0) Gecko/20100101 Firefox/5.0',\
              'Mozilla/5.0 (Windows NT 5.2) AppleWebKit/534.30 (KHTML, like Gecko) Chrome/12.0.742.122 Safari/534.30']

proxy_list=['115.214.10.108:8118','61.191.173.31:808',\
            '218.4.101.130:83']

def write_xlsx():
    w=xlsxwriter.Workbook(sys.path[0]+'\\'+keywords+'.xlsx')
    ws=w.add_worksheet('data')
    rowformat = w.add_format()
    rowformat.set_font_size(10) 
    rowformat.set_font_name('Microsoft yahei')
    rowformat.set_align('vcenter')
    rowformat.set_text_wrap() # 设置自动换行
    ws.set_column('A:A',60)
    ws.set_column('B:B',90)
    for i in xrange(len(all_que_list)):
        try:
            ws.set_row(i,20,rowformat)
            ws.write_string(i,0,all_que_list[i])
            ws.write_string(i,1,all_ans_list[i])
        except Exception,e:
            print'write_xlsx:',e
    w.close() 


class QSBK(threading.Thread):

    #初始化方法，定义一些变量
    def __init__(self,threadID,page_start,page_end):
        threading.Thread.__init__(self)
        self.threadID=threadID
        self.page_start=page_start
        self.page_end=page_end
        self.que_list=[]
        self.ans_list=[]
        
    def str_replace(self,old_str):
        new_str=re.sub('http://[a-zA-Z0-9./_=&?%]+','',old_str)
        new_str=re.sub('www.[a-zA-Z0-9./_=&?%]+','',new_str)
        new_str=re.sub('refid[0-9=]+','',new_str)
        new_str=re.sub('REFID[0-9=]+','',new_str)
        new_str=re.sub('<a.*?</a>','',new_str)
        new_str=new_str.replace('——','至')
        new_str=new_str.replace('℃','度')
      
        for split in split_list:
            #print split
            new_str=new_str.replace(split,',')
        for str1 in str_list:
            new_str=new_str.replace(str1,'')
        return new_str
    
    def getPage(self):
        try:
            header=random.choice(headers_list)
            #header=headers_list[0]
            proxy=random.choice(proxy_list)
            dict_header={'User-Agent':header}
            '''
            print proxy
            dict_proxy={'http':proxy}
            proxy_support=urllib2.ProxyHandler(dict_proxy)
            opener = urllib2.build_opener(proxy_support,urllib2.HTTPHandler)
            urllib2.install_opener(opener)
            '''
            request = urllib2.Request(self.url,headers=dict_header)# 随机取user-agent
            #利用urlopen获取页面代码
            response = urllib2.urlopen(request)
            #将页面转化为UTF-8编码
            pageCode = response.read().decode('utf-8')
            #time.sleep(random.uniform(0,1))#随机休眠1-3秒
            return pageCode
        except urllib2.URLError, e:
            if hasattr(e,"reason"):
                print u"连接新浪爱问失败,错误原因",e.reason
                return None
 
    def getURL(self):
        url_list=[]
        pageCode = self.getPage()
        if not pageCode:
            print "页面加载失败...."
            return None
        pattern = re.compile('<p class="title".*?<a href="(.*?)" target="_blank">',re.S)
        items = re.findall(pattern,pageCode)
        for item in items:
            #print item
            url_list.append(item)    
        return url_list
        
        
    def getPageItems(self):
        ans_doc_list=[]
        short_ans=''
        pageCode = self.getPage()
        if not pageCode:
            print "页面加载失败...."
            return None
        title_pattern=re.compile(u'<title>(.*?)- 爱问知识人</title>')
        title=re.findall(title_pattern,pageCode)
        #print title[0]
        question=self.str_replace(title[0])
        self.que_list.append(question)
        answer_pattern=re.compile('<span><pre>(.*?)</pre></span>',re.S)
        items = re.findall(answer_pattern,pageCode)
        for item in items:
            #print item
            ans_doc_list.append(item)    
        #print answer_list
        ans_doc_list.sort(key=lambda x:len(x))
        ans_doc_list.reverse()
        k=0
        while(k<len(ans_doc_list)):
            answer_length=len(ans_doc_list[k])
            if 60<answer_length<600:
                short_ans=ans_doc_list[k]
            elif answer_length<600:
                if not ans_doc_list[k].strip():
                    if k-1<0:
                        short_ans=ans_doc_list[0]
                    else:
                        short_ans=ans_doc_list[k-1]
                else:
                    short_ans=ans_doc_list[k]
            else:
                short_ans=ans_doc_list[-1]
            k=k+1
        short_ans=re.sub('<div.*?target="_blank">','',short_ans)
        short_ans=re.sub('<div.*?</div>+','',short_ans)
        short_ans=re.sub('<.*?>+','',short_ans)
        short_ans=self.str_replace(short_ans)
        self.ans_list.append(short_ans)


    def run(self):
        print u'爬虫%d号正在加载中...\n'%self.threadID
        for i in range(self.page_start,self.page_end):
            self.url ='http://iask.sina.com.cn/search?searchWord='+keywords+'&page='+str(i)
            url_list=self.getURL()
            if url_list==None:return 0
            if not url_list:
                print 'url_list is empty'
                time.sleep(300)
                return -1
            for b_url in url_list:
                self.url='http://iask.sina.com.cn'+b_url
                self.getPageItems() 
            time.sleep(random.uniform(0,1))#随机休眠0-5秒
        print u'爬虫%d号结束工作...\n'%self.threadID
        


if __name__=='__main__':
    
    t_num=5#线程数
    threads=[]
    all_que_list=[]
    all_ans_list=[]
    page_num=(end_num-start_num)/t_num
    for i in range(t_num):
        page_start=i*page_num+start_num
        page_end=(i+1)*page_num+start_num
        if i==t_num-1:
            print page_start,end_num+1
            thread1=QSBK(i+1,page_start,end_num+1)
        else:
            print page_start,page_end
            thread1=QSBK(i+1,page_start,page_end)
        threads.append(thread1)
    for t in threads:
        t.setDaemon(True)
        t.start()
        time.sleep(5)
    print u'爬虫正在收集数据...'
    for t in threads:
        t.join()
    print u'数据整理中...'
    for j in range(t_num):
        all_que_list=all_que_list+threads[j].que_list
        all_ans_list=all_ans_list+threads[j].ans_list
    print u'写入数据到excel表格...'
    write_xlsx()
    print u'测试结束'
    










    
