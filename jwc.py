# -*- coding:utf-8 -*-
# author:YangLilin
# Github:https://github.com/lilin5819
# GPL/BSD
import requests
import string
import sys,os
import xlwt
from bs4 import BeautifulSoup
from PIL import Image

class Spider(object):
    Host='jwc.xhu.edu.cn'
    charset=''
    usr_name=''
    usr_id=''
    usr_pass=''
    __VIEWSTATE=''
    __EVENTTARGET=''
    __EVENTARGUMENT=''

    header = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.8',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded',
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36',
    }

    login_post={
        '__VIEWSTATE':'',
        'txtUserName':'',
        'TextBox2':'',
        'txtSecretCode':'',
        'RadioButtonList1':'',
        'Button1':'',
        'lbLanguage':'',
        'hidPdrs':'',
        'hidsc':'',
    }


    def __init__(self,website):
        if website:
            self.Host=website
        self.base_url='http://'+self.Host
        self.login_url=self.base_url+'/default2.aspx'
        self.checkcode_url=self.base_url+'/CheckCode.aspx'
        self.xs_url=self.base_url+'/xs_main.aspx'
        self.kb_url=self.base_url+'/xskbcx.aspx'
        self.cj_url=self.base_url+'/xscjcx.aspx'


        self.session=requests.session()
        r=self.session.get(self.base_url)
        self.charset=r.encoding
        soup=BeautifulSoup(r.content,'lxml')
        self.__VIEWSTATE=soup.find('input',{'name':'__VIEWSTATE'})['value']
        self.login_post['__VIEWSTATE']=self.__VIEWSTATE

    def check_code(self):
        r=self.session.get(self.checkcode_url,stream=True)
        img=r.content
        with open('checkcode.jpg','wb') as jpg:
            jpg.write(img)
        jpg.close
        jpg=Image.open('checkcode.jpg')
        jpg.show()
        check_code=raw_input(u'请输入验证码:')
        return check_code

    def login(self,uid,upass):
        self.usr_id,self.usr_pass=uid,upass
        self.login_post['txtUserName']=self.usr_id
        self.login_post['TextBox2']=self.usr_pass
        self.login_post['txtSecretCode']=self.check_code()
        self.login_post['Button1']=u'登陆'.encode(self.charset)
        self.login_post['RadioButtonList1']=u'学生'.encode(self.charset)
        r=self.session.post(self.login_url,data=self.login_post,headers=self.header)
        self.header['Referer']=r.request.url
        print(r.request.url)

        soup=BeautifulSoup(r.content,'lxml')
        try:
            self.usr_name=soup.find('span',id='xhxm').get_text()[:-2]
            self.get_stat(soup)
        except:
            print(u'登陆失败，清检查学号、密码、验证码!!!')
            sys.exit()

        return r

    def get_stat(self,soup):
        self.__VIEWSTATE=soup.find('input',{'name':'__VIEWSTATE'})['value']
        self.__EVENTTARGET=soup.find('input',{'name':'__EVENTTARGET'})['value']
        self.__EVENTARGUMENT=soup.find('input',{'name':'__EVENTARGUMENT'})['value']

    def get_cj(self):
        payload={
            'xh':self.usr_id,
            'xm':self.usr_name.encode(self.charset),
            'gnmkdm':'N121605',
        }

        cj_post={
            '__EVENTTARGET': self.__EVENTTARGET,
            '__EVENTARGUMENT': self.__EVENTARGUMENT,
            '__VIEWSTATE':self.__VIEWSTATE,
            'hidLanguage': '',
            'ddlXN': '',
            'ddlXQ': '',
            'ddl_kcxz': '',
            'btn_zcj': u'历年成绩'.encode(self.charset),
        }

        r=self.session.get(self.cj_url,params=payload,headers=self.header)
        self.header['Referer']=r.request.url
        print('get cj')
        print(r.request.url)
        soup=BeautifulSoup(r.content,'lxml')
        self.get_stat(soup)
        cj_post['__VIEWSTATE']=self.__VIEWSTATE

        r=self.session.post(self.cj_url,params=payload,headers=self.header,data=cj_post)
        self.header['Referer']=r.request.url
        print('post cj')
        print(r.request.url)
        soup=BeautifulSoup(r.content,'lxml')
        self.get_stat(soup)
        usr_info=soup.find('table',{'class':'formlist'})
        class_info=soup.find('table',{'class':'datelist'})
        zxf=0.0                  #总学分
        zjd=0.0                  #总绩点
        m,n=0,0
        book=xlwt.Workbook()        #开辟一个xls
        sheet=book.add_sheet('sheet1')  #新建一页表格
        for line in class_info.find_all('tr'):
            n,xf,jd=0,0,0
            for arg in line.find_all('td'):
                s=arg.get_text().strip()
                if s and m and n==6:    #学分列
                    xf=string.atof(s)
                    sheet.write(m,n,xf)
                    zxf+=xf
                elif s and m and n==7:  #绩点列
                    jd=string.atof(s)
                    sheet.write(m,n,jd)
                    zjd+=jd
                else:
                    sheet.write(m,n,s)
                    if n==14 and s:    #减出重修多算的分
                        zxf-=xf
                n+=1
            m+=1
        print(u'总学分:%f'%zxf)
        print(u'总绩点:%f'%zjd)
        sheet.write(m,6,zxf)
        sheet.write(m,7,zjd)
        filename='%s.xls'%self.usr_name
        try:
            book.save(filename)
        except IOError:
            print(u'保存文件失败: %s ！'%filename)
            print(u'检查是否此文件被其他软件打开,关闭它,并重新查询成绩!')
            sys.exit()
        print(u'已生成：%s'%filename)

if __name__ == '__main__':
    if os.name == 'nt':
        reload(sys)
        sys.setdefaultencoding('gbk')
    spider=Spider('jwc.xhu.edu.cn')
    if len(sys.argv)<3:
        print(u'未输入参数!清重新输入学号和密码!!!!!')
        xh=raw_input(u'请输入学号:')
        mm=raw_input(u'请输入密码:')
        spider.login(xh,mm)
    else:
        spider.login(sys.argv[1],sys.argv[2])
    if spider.usr_name:
        print(u'登陆成功!')
        print(u'姓名:'+spider.usr_name)
        print(u'学号:'+spider.usr_id)
    else:
        print(u'登陆失败!')
        exit()
    spider.get_cj()
