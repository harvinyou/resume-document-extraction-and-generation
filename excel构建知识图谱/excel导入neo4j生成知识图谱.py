import os
import xlrd
import datetime
import time
import re
import pandas as pd
import unicodedata
from py2neo import Node, Graph, Relationship, NodeMatcher
graph = Graph('http://localhost:7474', username='neo4j', password='123456')
graph.delete_all() # 清除neo4j里面的所有数据
matcher = NodeMatcher(graph)
path ='验证集.xlsx'
xl = xlrd.open_workbook(path)
print("开始转换，请稍等······ ")
# 读取第一个表
table = xl.sheet_by_name(xl.sheet_names()[0])
nrows = table.nrows
firstrow = table.row_values(0)
listonce=[]
flag=0
for i in range(nrows):  # 循环逐行打印
    if i > 0:
        if table.row_values(i)[0]!='':
            node_1 = Node('姓名', name=str(table.row_values(i)[0]))#插入节点
            graph.create(node_1)  # 插入节点
            name_1 = matcher.match('姓名').where("_.name=" + "'" + str(table.row_values(i)[0]) + "'").first()
        if table.row_values(i)[0]=='':
            node_1 = Node('姓名', name=str(table.row_values(i)[18]))  # 插入节点
            graph.create(node_1)#插入节点
            name_1 = matcher.match('姓名').where("_.name=" + "'" + str(table.row_values(i)[18]) + "'").first()
        j=0
        k=0
        for content in table.row_values(i):
            if '[' in content and content!='':
                content.replace(" ", "")
                content=content[1:]
                content = content[:-1]
                content=content.split(',')
                if type(content)==list:
                    k = 0
                    for contenti in content:
                        contenti=contenti[1:]
                        contenti =contenti[:-1]
                        if k!=0:
                            contenti = contenti[1:]
                        # 判断是否存在该节点
                        unicodedata.normalize('NFKC', contenti)
                        flag=0
                        for temp in listonce:
                            if contenti!=''and contenti==temp:#尾实体重复
                                flag=1
                                break#尾实体重复
                        if flag==1:
                            name_2 = matcher.match('属性值').where("_.name=" + "'" + contenti + "'").first()
                        if flag==0 and contenti!='':#尾实体不重复加节点
                            listonce.append(contenti)
                            print(contenti,i,"----------列表字段")
                            node_2 = Node('属性值', name=contenti)#插入节点
                            graph.create(node_2)  # 插入节点
                            name_2 = matcher.match('属性值').where("_.name=" + "'" + contenti + "'").first()

                        relationship = Relationship(name_1, firstrow[j], name_2)
                        if relationship!=None:
                            graph.create(relationship)
                        k=k+1


            elif content!='':
                node_2 = Node('属性值', name=content)
                #判断是否存在该节点
                flag = 0
                for temp in listonce:
                    if content == temp:
                        flag = 1
                        break
                if flag == 0:
                    listonce.append(content)
                    print(content,i,"——————————————普通字段")
                    node_2 = Node('属性值', name=content)  # 插入节点
                    graph.create(node_2)  # 插入节点
                node_2 = Node('属性值', name=content)
                name_2 = matcher.match('属性值').where("_.name=" + "'" + content + "'").first()#插入节点
                relationship = Relationship(name_1, firstrow[j], name_2)
                graph.create(relationship)
            j=j+1




