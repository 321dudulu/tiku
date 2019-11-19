import xlrd
import easygui as eg
import  random
#定义一个类，提供相应的方法
class Titu():
    #初始化函数
    def __init__(self):
        print('第一步，初始化')
        self.data=None
        self.tables=None
        self.sigal=None
        self.mutip=None
        self.judge=None

        self.inputpath='20191113.xls'
        self.outputpath=''
        self.savemasknum=''
        self.signalnum=0
        self.multinum=0
        self.judgenum=0
        self.temp=[]

        self.sigal_rand_rows=[]
        self.multi_rand_rows=[]
        self.judge_rand_rows=[]

        self.sigal_answer=[]
        self.mult_answer=[]
        self.judge_answer=[]

        self.sigal_consist = 4
        self.mutip_consist = 4
        self.judge_consist = 4


    #获取分数，题数等相关信息
    def getinformation(self):
        print('第二步，获取文件信息')
        self.data=xlrd.open_workbook(self.inputpath)
        print('文件读取成功')
        self.tables=self.data.sheets()
        self.signalnum=eg.integerbox(msg='请输入单选题数量',title='输入信息',default='0')
        self.sigal_rand_rows=[0]*self.signalnum
        self.multinum=eg.integerbox(msg='请输入多选题数量',title='输入信息',default='0')
        self.multi_rand_rows=[0]*self.signalnum
        self.judgenum=eg.integerbox(msg='请输入判断题数量',title='输入信息',default='0')
        self.judge_rand_rows=[0]*self.multinum



    #读取相应的题号到一个内存中
    def getrandomnum(self):
        print('第三步，获取随机号')
        #提取单选题
        self.sigal=self.data.sheet_by_name('单选题')

        for i in range(self.signalnum):
            self.sigal_rand_rows[i]=random.randint(0,self.sigal.nrows)


        # 提取多选题
        self.mutip = self.data.sheet_by_name('多选题')

        for i in range(self.multinum):
            self.multi_rand_rows[i] = random.randint(0, self.mutip.nrows)

        # 提取判断题
        self.judge = self.data.sheet_by_name('判断题')

        for i in range(self.judgenum):
            self.judge_rand_rows[i] = random.randint(0, self.judge.nrows)
        print(self.sigal_rand_rows)

    def display(self):
        #单选题
        print('第四步，开始做题')
        c=self.sigal_consist
        result=[]
        tem_result=[]
        eg.msgbox(msg='首先是选择题，是否继续',title='infomation')
        tem_result=['']*self.signalnum
        for i in range(self.signalnum):
            temp_row=self.sigal.row(self.sigal_rand_rows[i])
            #获得答案
            self.sigal_answer=['']*self.signalnum
            self.sigal_answer[i]=temp_row[self.sigal_consist]
            #显示数据并或结果
            tem_result[i]=eg.buttonbox(msg=str(temp_row[c-2])+'\n\n\n',choices=(str(temp_row[c+1]),str(temp_row[c+2]),str(temp_row[c+3]), str(temp_row[c+4])))
            print(tem_result[i])
        result.append(tem_result)

        #多选题
        eg.msgbox(msg='接下来是多选题，是否继续', title='infomation')
        tem_result=['']*self.multinum
        for i in range(self.multinum):
            temp_row = self.mutip.row(self.multi_rand_rows[i])
            # 获得答案
            self.mult_answer=['']*self.multinum
            self.mult_answer[i] = temp_row[self.mutip_consist]
            # 显示数据并或结果
            tem_result[i] = eg.choicebox(msg=str(temp_row[c - 2]) + '\n\n\n', choices=(
            str(temp_row[c + 1]), str(temp_row[c + 2]), str(temp_row[c + 3]), str(temp_row[c + 4])))
        result.append(tem_result)

        #判断题
        eg.msgbox(msg='最后是判断题，是否继续', title='infomation')
        tem_result=['']*self.judgenum
        for i in range(self.judgenum):
            temp_row = self.judge.row(self.sigal_rand_rows[i])
            # 获得答案
            self.judge_answer=['']*self.judgenum
            self.judge_answer[i] = temp_row[self.judge_consist]
            # 显示数据并或结果
            tem_result[i] = eg.buttonbox(msg=str(temp_row[c - 2]) + '\n\n\n', choices=(
           '正确','错误'))
        result.append(tem_result)

    #保存错误号
    def savenumber(self):
        pass


if __name__=='__main__':
    t=Titu()
    t.getinformation()
    t.getrandomnum()
    t.display()