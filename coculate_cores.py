import xlrd
import xlwt
import easygui

class Coclulate_utils():
    #对象属性
    paper_data_path='./coculate.txt'

    #对考试成绩每个模块求和
    def sum(self,st):
        str_list=st.split(',')
        len_list=len(str_list)
        if len_list!=5:
            print('输入错误，请重新输入\n')
            return None
        sum=0
        for i in range(len_list):
            temp=int(str_list[i])
            sum+=temp
        print(sum)
        return sum

    # 卷子成绩统计到该文件中
    #接受数据，并将其写入到对应的文件中，格式例如 张三，part1,part2,part3,part4,part5,toll_scores
    def colulate_core_paper(self):
        flag=True

        print("请按顺序输入每个部分的数据\n")
        print('格式为part1,part2,part3,part4,part5')
        sheet=self.read('course05.xls')
        col_name=sheet.col_values(2)
        col_number=sheet.col_values(1)
        length=len(col_name)
        for i in range(length)[1:length]:
            f = open(self.paper_data_path, 'a')
            str1=input("输入"+col_name[i]+'  '+col_number[i]+'信息\n')
            sum_result=self.sum(str1)
            flag=False
            if sum_result is None:
                flag=True
            while sum_result is None:
                str1 = input("输入" + col_name[i] + '  ' + col_number[i] + '信息\n')
                sum_result = sum(str1)

            wr_content=str(col_name[i])+','+str(col_number[i])+','+str1+','+str(sum_result)+'\n'
            f.writelines(wr_content)
            f.close()

    #读取excel数据
    def read(self,name):
        workbook=xlrd.open_workbook(name)
        sheet=workbook.sheet_by_index(0)
        return sheet

    #设置写入数据的样式
    def set_style(self,font_name,font_height,font_bold):
        font=xlwt.Font()
        style=xlwt.XFStyle
        font.name=font_name
        font.bold=font_bold
        font.height=font_height
        font.colour_index=4
        style.font=font
        return  style


#具体写入时的格式
#sheet1.write(0, i,row0[i], set_style( &apos;Times New Roman&apos;,220, True))
#表的属性 ['序号', '学号', '姓名', '平时作业成绩', '实验课表现', '大作业成绩', '理论课表现', '期末考试', '备注']
if __name__=='__main__':
    c=Coclulate_utils()
    #c.read('course05.xls')
    c.colulate_core_paper()
