#coding=utf-8

#初次运行请在系统自带的cmd或powershell中输入如下代码以安装必要依赖
#pip install --upgrade pip
#pip install pandas openpyxl matplotlib


#from symbol import try_stmt
#from ast import Try
#from symbol import try_stmt
#from curses import noecho
from tempfile import TemporaryDirectory
import numpy
import pandas
import matplotlib.pyplot
import os
import sys
import shutil
import subprocess
import copy
matplotlib.pyplot.switch_backend('agg')


'''
#此时tt是工作路径下所需的.xlsx的文件名
pd=pandas

# 读取数据表格
file_path = tt # 请替换为您的文件路径
sheet_name = 'Sheet1'  # 请替换为您的表格名称

# 读取Excel表格
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 删除在18题或19题中选择了"携带的所有手机在使用完毕后都需要由自己吃掉，不能浪费"的行
df_filtered = df[~((df['18、您所就读（或曾就读）的高中，对于手机，笔记本电脑，平板电脑等智能电子设备的态度如何？'] == "携带的所有手机在使用完毕后都需要由自己吃掉，不能浪费") | 
                   (df['19、您所就读的学校，对于手机，笔记本电脑，平板电脑等智能电子设备的态度如何？'] == "携带的所有手机在使用完毕后都需要由自己吃掉，不能浪费"))]

# 删除在3题中选择了"球外"的行
df_filtered = df_filtered[df_filtered['3、您的出生地大致来说在'] != "球外"]

# 用户自定义筛选条件
# 例如，假设用户希望筛选 "Q5" 列中选择了 "是" 的问卷
#user_defined_filter = df_filtered['Q5'] == "是"  # 替换条件

# 应用用户自定义筛选条件
#df_result = df_filtered[user_defined_filter]

# 将结果保存到新Excel文件

        
output_file_path = './mcache/'+-str(index_of_excel)+'.xlsx'
df_filtered.to_excel(output_file_path, index=True)

print(f"处理完成，结果已保存至：{output_file_path}")
os.system(output_file_path)


'''

class Otherthing_:
    def cleanwindow(self):#清屏
       
        sys.stdout.write("\033[2J\033[H")
        sys.stdout.flush()
    def new_xl_idx(self):
        with open("./info","r") as i:
            index_of_excel=int(i.readline())
            index_of_temp=int(i.readline())
            global now_xl
            now_xl=i.readline()
        with open("./info","w") as i:
            i.write(str(index_of_excel+1)+'\n')
            i.write(str(index_of_temp)+'\n')
            i.write(now_xl)
        return index_of_excel

    def new_tp_idx(self):
        with open("./info","r") as i:
            index_of_excel=int(i.readline())
            index_of_temp=int(i.readline())
            global now_xl
            now_xl=i.readline()
        with open("./info","w") as i:
            i.write(str(index_of_excel)+'\n')
            i.write(str(index_of_temp+1)+'\n')
            i.write(now_xl)
        return index_of_temp    
    
    def change_nowxl(self,newone):
        with open("./info","r") as i:
            index_of_excel=int(i.readline())
            index_of_temp=int(i.readline())
            global now_xl
            now_xl=newone
        with open("./info","w") as i:
            i.write(str(index_of_excel)+'\n')
            i.write(str(index_of_temp)+'\n')
            i.write(newone)

    def get_xl_idx(self):
        with open("./info","r") as i:
            index_of_excel=int(i.readline())
            index_of_temp=int(i.readline())
            global now_xl
            now_xl=i.readline()
        with open("./info","w") as i:
            i.write(str(index_of_excel+1)+'\n')
            i.write(str(index_of_temp)+'\n')
            i.write(now_xl)
        return index_of_excel-1

    def get_tp_idx(self):
        with open("./info","r") as i:
            index_of_excel=int(i.readline())
            index_of_temp=int(i.readline())
            global now_xl
            now_xl=i.readline()
        return index_of_temp-1    
    
    def get_nowxl(self):
        with open("./info","r") as i:
            index_of_excel=int(i.readline())
            index_of_temp=int(i.readline())
            global now_xl
            now_xl=i.readline()
    def new_xl(self,frame,idx=False):
        frame.to_excel('./work/'+str(self.new_xl_idx())+".xlsx",index=idx)
    def new_tp(self,frame,idx=False):
        frame.to_excel('./mcache/'+str(self.new_tp_idx())+".xlsx",index=idx) 
        
    def oif_in(self):
        with open("./oif","r") as i:
           global temp_open_on
           ii=i.readlines()
           for iii in range(len(ii)):
               ii[iii]=int(ii[iii])
           temp_open_on=ii[0]
    def oif_out(self,idx,val)   :
        with open("./oif","r") as i:
           global temp_open_on
           ii=i.readlines()
        with open("./oif","w") as i:
           for iii in range(len(ii)):
               i.write(ii[iii]if iii!=idx else val)
    def format_xl(self):
        global now_frame
        lab=now_frame.columns
        for ii in range(len(lab)):
            try:
                now_frame[lab[ii]].astype('int64')
            except:
                try:
                    now_frame[lab[ii]].astype('float64')
                except:
                    try:
                        now_frame[lab[ii]].astype('category')
                    except:
                        pass

main_menu='''1   操作当前表格
2   切换当前表格
3   将刚才生成的临时表格保存
33  -将临时表格保存然后切换为当前表格
4   查看当前表格
5   其它设置
6   退出
'''
oper_menu='''1   按条件筛选
2   按条件排除
22  -去除搞笑选项
3   替换选项
55  -按默认规则替换字符为数字
7   删除问题列
77  -删除非数问题
777 --删除选项全部重复的列
7777---去除搞笑选项、按默认规则替换字符为数字、删除非数问题和选项全部重复的列
9   数据处理
    1    协方差
    2    相关系数
4   (测试)导出替换
5   (测试)导入替换
6   (测试)手动格式化数据类型
8   (测试)输出所有数据类型
'''
now_xl=''

text_cache=''

temp_open_on=1

now_frame=pandas.DataFrame({})



#1

def oper(x=0,y=[]):
    O.cleanwindow()
    print(oper_menu)
    if x==0:
        temp=int(input('\n>>'))
    else:
        temp=x
    global now_frame 
    global text_cache
    global temp_open_on
    
    label=now_frame.columns
    
    if temp==1:
        O.cleanwindow()
        print('\n\n\n')
        for ii in range(len(label)):
            print(ii+1,'   ',label[ii])
        temp=int(input('\n选择问题>>'))-1
        answerlist=[]
        now_q=label[temp]
        ttt=now_frame[now_q]
        def almk(v,al=answerlist):
            f=True
            for i in al:
                if i==v:
                    f=False
                    break
            if f:
                al+=[v]
        ttt.apply(almk)
        O.cleanwindow()
        for ii in range(len(answerlist)):
            print(ii+1,'   ',answerlist[ii])
        
        alndaser=[]
        temp=int(input('\n选择要保留的答案>>'))-1
        alndaser+=[answerlist[temp]]
        del answerlist[temp]
        while temp>=0 and len(answerlist)!=0:
            O.cleanwindow()
            for ii in range(len(answerlist)):
                print(ii+1,'   ',answerlist[ii])
            temp=int(input('已选择'+str(alndaser)+'\n输入0结束选择\n\n或继续选择要保留的答案>>'))-1
            if temp<0:
                break
            alndaser+=[answerlist[temp]]
            del answerlist[temp]
        tttt=now_frame[now_q]==alndaser[0]
        ttttt=alndaser[:]
        del alndaser[0]
        if alndaser!=[]:
            for ii in alndaser:
                tttt=tttt|(now_frame[now_q]==ii)
        now_frame=now_frame[tttt]
        O.format_xl()
        O.new_tp(now_frame)
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx')) 

    elif temp==2:
        O.cleanwindow()
        print('\n\n\n')
        for ii in range(len(label)):
            print(ii+1,'   ',label[ii])
        temp=int(input('\n选择问题>>'))-1
        answerlist=[]
        now_q=label[temp]
        ttt=now_frame[now_q]
        def almk(v,al=answerlist):
            f=True
            for i in al:
                if i==v:
                    f=False
                    break
            if f:
                al+=[v]
        ttt.apply(almk)
        O.cleanwindow()
        for ii in range(len(answerlist)):
            print(ii+1,'   ',answerlist[ii])
        
        alndaser=[]
        temp=int(input('\n选择要排除的答案>>'))-1
        alndaser+=[answerlist[temp]]
        del answerlist[temp]
        while temp>=0 and len(answerlist)!=0:
            O.cleanwindow()
            for ii in range(len(answerlist)):
                print(ii+1,'   ',answerlist[ii])
            temp=int(input('已选择'+str(alndaser)+'\n输入0结束选择\n\n或继续选择要排除的答案>>'))-1
            if temp<0:
                break
            alndaser+=[answerlist[temp]]
            del answerlist[temp]
        tttt=now_frame[now_q]==alndaser[0]
        ttttt=alndaser[:]
        del alndaser[0]
        if alndaser!=[]:
            for ii in alndaser:
                tttt=tttt|(now_frame[now_q]==ii)
        now_frame=now_frame[~tttt]
        O.format_xl()
        O.new_tp(now_frame)
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache='ok\n'
    
    elif temp==22:
        now_frame=now_frame[~((now_frame['18、您所就读（或曾就读）的高中，对于手机，笔记本电脑，平板电脑等智能电子设备的态度如何？'] == "携带的所有手机在使用完毕后都需要由自己吃掉，不能浪费") | 
                   (now_frame['19、您所就读的学校，对于手机，笔记本电脑，平板电脑等智能电子设备的态度如何？'] == "携带的所有手机在使用完毕后都需要由自己吃掉，不能浪费")|(now_frame['3、您的出生地大致来说在'] == "球外"))]
        O.format_xl()
        O.new_tp(now_frame)
        now_frame=pandas.read_excel('./mcache/'+str(O.get_tp_idx())+'.xlsx')
        ttttt='已去除'
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))  
        else:
            text_cache='ok\n'
    elif temp==3:
        O.cleanwindow()
        print('\n\n\n')
        for ii in range(len(label)):
            print(ii+1,'   ',label[ii])
        temp=int(input('\n选择问题>>'))-1
        answerlist=[]
        now_q=label[temp]
        ttt=now_frame[now_q]
        def almk(v,al=answerlist):
            f=True
            for i in al:
                if i==v:
                    f=False
                    break
            if f:
                al+=[v]
        ttt.apply(almk)
        O.cleanwindow()
        for ii in range(len(answerlist)):
            print(ii+1,'   ',answerlist[ii])
        
        alndaser=[]
        a1={}
        temp=int(input('\n选择要替换的答案>>'))-1
        
        t1=input('替换为>>')
        try:
            t1=int(t1)
        except:
            pass    
        alndaser+=[answerlist[temp]]
        a1[answerlist[temp]]=t1
        del answerlist[temp]
        
        while temp>=0 and len(answerlist)!=0:
            O.cleanwindow()
            for ii in range(len(answerlist)):
                print(ii+1,'   ',answerlist[ii])
            temp=int(input('已选择'+str(alndaser)+'\n输入0结束选择\n\n或继续选择要替换的答案>>'))-1
            if temp<0:
                break
            
                
            t1=input('替换为>>') 
            try:
                t1=int(t1)
            except:
                pass   
            alndaser+=[answerlist[temp]]
            
   
            a1[answerlist[temp]]=t1
            alndaser=alndaser[:]
            del answerlist[temp]
        now_frame.loc[now_frame[now_q]==alndaser[0],now_q]=a1[alndaser[0]]
        ttttt=alndaser[:]
        del alndaser[0]
        if alndaser!=[]:
            for ii in alndaser:
                now_frame.loc[now_frame[now_q]==ii,now_q]=a1[ii]
        #now_frame=now_frame[~tttt]
        O.format_xl()
        O.new_tp(now_frame)
        now_frame=pandas.read_excel('./mcache/'+str(O.get_tp_idx())+'.xlsx')
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache='ok\n'
    elif temp==4:    
      va={}
      while temp==4:
        
        
        print('\n\n\n')
        O.cleanwindow()
        
        print(va)
        for ii in range(len(label)):
            print(ii+1,'   ',label[ii])
        temp=int(input('\n选择问题>>'))-1
        answerlist=[]
        now_q=label[temp]
        ttt=now_frame[now_q]
        def almk(v,al=answerlist):
            f=True
            for i in al:
                if i==v:
                    f=False
                    break
            if f:
                al+=[v]
        ttt.apply(almk)
        O.cleanwindow()
        for ii in range(len(answerlist)):
            print(ii+1,'   ',answerlist[ii])
        
        alndaser=[]
        a1={}
        temp=int(input('\n选择要替换的答案>>'))-1
        if temp<0:
            temp=4
            continue
        t1=input('替换为>>')
        try:
            t1=int(t1)
        except:
            pass    
        alndaser+=[answerlist[temp]]
        a1[answerlist[temp]]=t1
        del answerlist[temp]
        
        while temp>=0 and len(answerlist)!=0:
            O.cleanwindow()
            for ii in range(len(answerlist)):
                print(ii+1,'   ',answerlist[ii])
            temp=int(input('已选择'+str(alndaser)+'\n输入0结束选择\n\n或继续选择要替换的答案>>'))-1
            if temp<0:
                break
            
                
            t1=input('替换为>>') 
            try:
                t1=int(t1)
            except:
                pass   
            alndaser+=[answerlist[temp]]
            
   
            a1[answerlist[temp]]=t1
            alndaser=alndaser[:]
            del answerlist[temp]
        #now_frame.loc[now_frame[now_q]==alndaser[0],now_q]=a1[alndaser[0]]
        ttttt=alndaser[:]
        va[now_q]=a1
        va=copy.deepcopy(va)
        #del alndaser[0]
        #if alndaser!=[]:
        #    for ii in alndaser:
        #        now_frame.loc[now_frame[now_q]==ii,now_q]=a1[ii]
        #now_frame=now_frame[~tttt]
        #O.new_tp(now_frame)
        #if temp_open_on:
        #    O.cleanwindow()
        #    print(ttttt,'\n正在查看...')
        #    os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        temp=int(input('输入4继续>>'))
        text_cache=str(va)
    elif temp==5:
        O.cleanwindow()
        temp=dict(input('导入替换规则>>'))
        for ii in temp: 
            for iii in temp[ii]:
                now_frame.loc[now_frame[ii]==iii,ii]=temp[ii][iii]
        #now_frame=now_frame[~tttt]
        O.format_xl()
        O.new_tp(now_frame)
        now_frame=pandas.read_excel('./mcache/'+str(O.get_tp_idx())+'.xlsx')
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache='ok\n'

    elif temp==55:
        temp={'2、您目前处于的年龄段': {'18~22': 19, '15~18': 17, '30岁以上': 40, '22~30': 26, '15岁以下': 13}, '4、您高中所在的城镇或村庄大约规模多大？': {'千人以下': 1, '小型城镇': 2, ' 中型城市': 3, '省会城市': 4, '北上广深': 5,'(跳过)': 0},'5、您学校所在的城镇或村庄大约规模多大？': {'千人以下': 1, '小型城镇': 2, ' 中型城市': 3, '省会城市': 4, '北上广深': 5,'(跳过)': 0}, '9、您是否有过通过自己和其他人共同的努力，齐心协力完成一项富有挑战性的工作的经历？': {'有': 1, '不记得有过了': 0}, '6、您对教育现代化怎么看？': {'看好': 1, '不清楚': 0, '不看好': -1},'24、您所就读的学校是否会使用app布置教学任务、收发作业、或考勤？': {'从不': 0, '有时': 1, '(跳过)': 0},'7、您对网络和电子产品是否熟悉？': {'十分熟悉': 3, '比较熟悉': 2, '不太熟悉': 1, '不熟悉': 0}, '22、您所就读（或曾就读）的高中是否会使用app布置教学任务、收发作业、或考勤？': {'经常': 3, '有时': 2, '很少': 1, '从不': 0, '(跳过)': 0}, '18、您所就读（或曾就读）的高中，对于手机，笔记本电脑，平板电脑等智能电子设备的态度如何？': {'禁止携带进入学校，严格控制在家中的使用': 0, '原则上不允许携带进入学校，但管控较松': 1, '学生可以自由携带': 2, '鼓励学生携带，或主动为学生提供': 3, '(跳过)': 0},'1、您的性别：':{'男':-1,'女':1},'8、网络上的各种学习资源和软件对你有帮助吗？': {'很有帮助': 2, '有一定帮助': 1, '(跳过)': 0, '没有帮助': -1}, '12、您认为现在的教育需要全球化、与世界接轨吗？': {'非常需要': 2, '需要': 1, '不太需要': 0, '完全不需要': -1}, '14、您认为教育全球化和现代化是否加剧了教育资源的不平等分配？': {'反而促进了资源平等': -1, '没有影响': 0, '一定程度上加剧了': 1, '是的，加剧了': 2}, '15、您认为可预期的将来新的教学模式是否会取代传统教育模式？': {'会几乎完全取代': 2, '会一定程度上取代': 1, '会少量的取代': 0, '不会取代': -1}, '19、您所就读的学校，对于手机，笔记本电脑，平板电脑等智能电子设备的态度如何？': {'( 跳过)': 0, '原则上不允许携带进入学校，但管控较松': 1}, '20、您认为你所就读（或曾就读）的高中对科技设备的使用是否提升了教学效果？': {'非常有效': 2, '有效': 1, '一般': 0, '(跳过)': 0, '不确定': 0, '无效': -1}, '21、您认为你所就读的学校对科技设备的使用是否提升了教学效果？': {'(跳过)': 0, '非常有效': 2, '有效': 1}, '23、您认为你所就读（或曾就读）的高中对app的使用是否有效提升了效率？': {'非常有效': 2, '有效': 1, '一般': 0, '(跳过)': 0, '无效': -1}, '25、您认为你所就读（或曾就读）的高中对app的使用是否有效提升了效率？': {'有效': 1, '(跳过)': 0}}
        for ii in temp:
            for iii in temp[ii]:
                now_frame.loc[now_frame[ii]==iii,ii]=temp[ii][iii]
        #now_frame=now_frame[~tttt]
        O.format_xl()
        O.new_tp(now_frame)
        now_frame=pandas.read_excel('./mcache/'+str(O.get_tp_idx())+'.xlsx')
        if temp_open_on:
            O.cleanwindow()
            print('\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache='ok\n'




    elif temp==6:
        O.format_xl()




    elif temp==7:
        for ii in range(len(label)):
            print(ii+1,'   ',label[ii])
        temp=int(input('\n选择问题>>'))-1
        now_q=label[temp]

        del now_frame[now_q]

        O.new_tp(now_frame)
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache='ok\n'


    elif temp==77:
        iii=0
        tt=[]
        O.format_xl()
        tttt=now_frame.dtypes
        for ii in tttt:
            print(ii)
            if str(ii)!='float64' and str(ii)!='int64':
                
                tt+=[now_frame.columns[iii]]
                tt=tt[:]
            iii+=1

        print(tt)
        for ii in tt:
            del now_frame[ii]
        O.new_tp(now_frame)
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache='ok\n'

    elif temp==777:
        ttttt=[]
        for temp in range(len(label)):
            answerlist=[]
            now_q=label[temp]
            ttt=now_frame[now_q]
            def almk(v,al=answerlist):
                f=True
                for i in al:
                    if i==v:
                        f=False
                        break
                if f:
                    al+=[v]
            ttt.apply(almk)
            if len(answerlist)==1:
                del now_frame[now_q]
                ttttt+=[now_q]
                ttttt=ttttt[:]
        O.new_tp(now_frame)
        if temp_open_on:
            O.cleanwindow()
            print(ttttt,'\n正在查看...')
            os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
        else:
            text_cache=str(ttttt)+'\nok\n'

    elif temp==7777:
        ttt=temp_open_on
        temp_open_on=0
        oper(22)
        oper(55)
        oper(77)
        oper(777)
        temp_open_on=ttt
    elif temp==8:
        iii=0
        for ii in now_frame.dtypes:
            print(now_frame.columns[iii],ii)
            iii+=1

    elif temp==9:
        O.cleanwindow()
        temp=int(input('选择>>'))
        if temp==1:
            now_frame=now_frame.cov()
            #print(now_frame.cov())
            O.new_tp(now_frame,idx=True)
            if temp_open_on:
                O.cleanwindow()
                print('\n正在查看...')
                os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
            else:
                text_cache='ok\n'
        elif temp==2:
            now_frame=now_frame.corr()
            #print(now_frame.cov())
            O.new_tp(now_frame,idx=True)
            if temp_open_on:
                O.cleanwindow()
                print('\n正在查看...')
                os.system(os.path.abspath('./mcache/'+str(O.get_tp_idx())+'.xlsx'))
            else:
                text_cache='ok\n'

    



def xl_init():
    O.cleanwindow()
    wk=True
    mc=True
    print('初始化中\n当前路径：',os.getcwd())
    allxl=[]
    for i in range(len(t)):
        tt=t[i]
        if tt[-5:]==".xlsx":
            allxl+=[tt][:]
        if tt=='work':
            wk=False
        if tt=='mcache':
            mc=False
    if len(allxl)==1:
        tt=allxl[0]
    else:
        for i in range(len(allxl)):
            print(i+1,'   ',allxl[i])
        tt=allxl[int(input('\n哪一份表格>>'))-1]
        
    with open("./info","w") as i:
        i.write('0\n0\n0')
    with open("./oif","w") as i:
        i.write('1')
    if wk:
        os.mkdir('work')
    if mc:
        os.mkdir('mcache')    
    
    global now_frame
    now_frame=pandas.read_excel(tt,index_col=0) 
    O.format_xl()
    O.new_xl(now_frame)
    O.change_nowxl("0.xlsx")
        
def o_init():
    O.get_nowxl()
    O.oif_in()
    global now_frame
    try:
        now_frame=pandas.read_excel('./work/'+now_xl)
        O.format_xl()
    except:
        ll=os.listdir("./work")
        O.cleanwindow()
        
        allxl=[]
        for i in range(len(ll)):
            tt=ll[i]
            if tt[-5:]==".xlsx":
                allxl+=[tt][:]

        if len(allxl)==1:
            tt=allxl[0]
        else:
            for i in range(len(allxl)):
                print(i+1,'   ',allxl[i])
            tt=allxl[int(input('\n哪一份表格>>'))-1]
        
        O.change_nowxl(tt)
        
        now_frame=pandas.read_excel('./work/'+tt)
        O.format_xl()
        





def the_main():
    O.cleanwindow()
    global text_cache
    global temp_open_on
    global now_frame
    print(text_cache+main_menu)
    text_cache=''
    temp=int(input('>>'))
    if temp==1:
        oper()


    elif temp==2:
        ll=os.listdir("./work")
        O.cleanwindow()
        allxl=[]
        for i in range(len(ll)):
            tt=ll[i]
            if tt[-5:]==".xlsx":
                allxl+=[tt][:]


        for i in range(len(allxl)):
            print(i+1,'   ',allxl[i])
        tt=allxl[int(input('\n哪一份表格>>'))-1]
        O.change_nowxl(tt)
        now_frame=pandas.read_excel('./work/'+tt)
        O.format_xl()
        text_cache="ok\n"


    elif temp==3:
        shutil.copy('./mcache/'+str(O.get_tp_idx())+'.xlsx','./work/'+O.new_xl_idx()+'.xlsx')
        text_cache='ok\n'


    elif temp==33:
        t=str(O.new_xl_idx())
        shutil.copy('./mcache/'+str(O.get_tp_idx())+'.xlsx','./work/'+t+'.xlsx')
        
        O.change_nowxl(t+'.xlsx')
        now_frame=pandas.read_excel('./work/'+t+'.xlsx')
        O.format_xl()
        text_cache='ok\n'
        


    elif temp==4:
        O.cleanwindow()
        print("正在查看...")
        os.system(os.path.abspath("./work/"+now_xl))
        
    elif temp==5:
        print(f'''1   临时文件{'不'if temp_open_on else ''}自动预览
-1   返回''')
        temp=int(input('>>'))
        if temp==1:
            O.oif_out(0,'0'if temp_open_on else '1' )
            #global temp_open_on
            temp_open_on=0 if temp_open_on else 1
        elif temp==-1:
            pass
        text_cache='ok\n'
        

    elif temp==6:
        exit()

t=os.listdir()
f=True
O=Otherthing_()
for i in range(len(t)):
    if t[i]=='info':
        f=False
if f:
    xl_init()
else:
    o_init()    

#如果工作路径中没有info文件就执行初始化    

  

while True:
    the_main()

#主循环    