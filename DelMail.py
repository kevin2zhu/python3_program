# -*- coding:utf-8 -*-
import imapclient
import os
import time



class DelMail():
    # 创建一个删除邮件的程序

    def __init__(self, server='imap.exmail.qq.com'):
        self.server = server

    def connet_server(self):
        # 连接邮件服务器

        try:
            self.mail = imapclient.IMAPClient(self.server)  # 连接邮箱服务器
            print('已连接上腾讯企业邮箱!')
            return self.mail
        except Exception as e:                              # 判断是否正常连接邮件服务器
            print(e,'\n邮箱服务器连接失败,请检查网络连接是否正确.\n')

    def login(self):
        # 登录邮件服务器

        while True:
            self.username = input('请输入邮箱账户: ')
            self.passwd = input('请输入邮箱密码: ')
            try:
                self.mail.login(self.username, self.passwd)  # 登录自己的邮箱
                print('\n登录邮箱成功!\n')
                break
            except imapclient.exceptions.LoginError as e:    # imapclient.error类重命名为e
                result = str(e)
                if 'email address' in result:                # 将错误码进行分析，如果带地址说明地址错误
                    print('\n邮箱账户格式错误,请检查后重试.\n')
                    time.sleep(1)
                elif 'Password' in result:
                    print('\n邮箱密码错误,请检查后重试.\n')      # 如果带password说明密码错误。
                    time.sleep(1)
                elif 'app password' in result:
                	print('\n请使用APP专用密码登录,一般在绑定的微信中可以查询的到.\n')
                	time.sleep(1)
                else:
                    print(e)

    def texts(self):
        # 文字说明
        print('重要!')
        print('请仔细阅读以下信息:')
        print('本程序默认删除收件箱中的所有邮件,如需指定查询条件删请看下例。')
        print('条件匹配只支持查询日期条件删除,格式:BEFORE/SINCE 01-May-2017。')
        print('BEFORE 01-May-2017(匹配2017年5月1日以前收到的邮件,不包括5月1号)。')
        print('SINCE 01-May-2017(匹配2017年5月1日之后收到的邮件,包括5月1号)。')
        print('经测试收件箱中的邮件被删除后,已删除文件夹的邮件也会清空,删除操作请慎重!')
        print('请一定备份好重要数据,并核对好要删除的邮件数目!')

    def decide(self):
        # 循环控制函数

        while True:
            mails = input('请核对需要删除的邮件数目是否正确?(yes/no)\n')
            if mails.lower() == ('yes' or 'y'):
                return self.marks
            elif mails.lower() == ('no' or 'n'):
                self.retopic()
                break
            else:
                print('请重新输入正确选项.\n')

    def retopic(self):
        #重写匹配条件

        while True:
            topic = input('\n请输入查询条件,eg:SINCE 01-May-2017:\n')
            try:
                self.mail.search(topic)
                self.marks = self.mail.search(topic)
                print('\n符合查询条件的邮件有%s封。' % len(self.marks))
                self.decide()
                return self.marks
            except Exception:
                print('\n条件匹配错误,请重试.')


    def delemail(self, topic='ALL'):
        # 删除邮件

        self.topic = topic
        self.mailnums = self.mail.select_folder('INBOX')
        print('\n邮箱中共有%s邮件。\n' % (self.mailnums[b'EXISTS']))

        while True:
            choince = input('默认删除收件箱中所有邮件,是否进行条件筛选?(yes/no)\n')
            if choince.lower() == ('yes' or 'y'):
                topic = input('请输入查询条件,eg:BEFORE 01-May-2017:\n')
                try:
                    self.marks = self.mail.search(topic)
                    print('\n符合查询条件的邮件有%s封。' % len(self.marks))
                    self.decide()
                    break
                except Exception:
                    print('\n条件匹配错误,请重试.')
                    self.retopic()
                    #self.decide()
                    break
            elif choince.lower() == ('no' or 'n'):
                topic = 'ALL'
                self.marks = self.mail.search(topic)
                print('\n需要删除的邮件有%s封.\n' % (self.mailnums[b'EXISTS']))
                break
            else:
                print('请重新输入正确选项.\n')

        while True:
            last = input('(慎重)是否进行邮件删除?(yes/no)\n')
            if last.lower() == ('yes' or 'y'):
                self.mail.delete_messages(self.marks)
                self.mail.expunge()
                self.mail.logout()
                print('邮件删除完毕,正在退出邮箱...')
                time.sleep(1)
                print('Bye,Bye\n')
                os.system('pause')
                break
            elif last.lower() == ('no' or 'n'):
                self.mail.logout()
                print('\n取消删除邮件,正在退出邮箱...')
                time.sleep(1)
                print('Bye,Bye\n')
                os.system('pause')
                break
            else:
                print('请重新输入正确选项.\n')




my_mail = DelMail()
if my_mail.connet_server():
    my_mail.login()
    my_mail.texts()
    time.sleep(3)
    my_mail.delemail()
else:
    os.system('pause')
