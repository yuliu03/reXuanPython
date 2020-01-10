#coding=gbk
import os

from openpyxl import load_workbook

def logic(input_list,expect_value,deviation):
    sum = 0
    for item in input_list:
        sum = sum + item
    if abs(sum-expect_value) <= deviation:
        print(input_list)

def isNotContain(start_pos, end_pos, core_pos):
    return (start_pos > core_pos and end_pos > core_pos or
    start_pos < core_pos and end_pos < core_pos)

#获取start_pos起始点到end_pos终点的数组
def getSegment(start_pos,end_pos,l):
    segment = list()
    #判断是否超出范围
    if end_pos >= len(l):
        return segment

    # 开始对待每个组合
    local_start_pos = start_pos
    local_end_pos = end_pos

    while local_start_pos <= local_end_pos:
        segment.append(l[local_start_pos])
        local_start_pos = local_start_pos + 1

    # 结束对待每个组合
    return segment

#获取所有的某长度的数组
def getAllSegment(s_length,l,s_pos,with_value):
    local_start_pos = 0
    local_end_pos = s_length-1

    allSegment = list()

    #特殊处理
    if s_length == 0:
         simpleList = [l[s_pos]]
         allSegment.append(simpleList)
         return allSegment

    # 通过localEnd的值判断是否还有组合 并且 去除不需要的数值
    while local_end_pos < len(l):
        if (local_start_pos > s_pos and local_end_pos > s_pos or local_start_pos < s_pos and local_end_pos < s_pos):
            segment =getSegment(local_start_pos,local_end_pos,l)
            if len(segment) > 0:
                #如果with_value为1，就包括s_pos的数值
                if with_value == 1:
                    segment.append(l[s_pos])
                allSegment.append(segment)

        local_end_pos = local_end_pos + 1
        local_start_pos = local_start_pos + 1

    return allSegment

#获取所有的elem对应的数组，并且组合
def getAllElemAllSegment(begin_length,max_length,l):
    data_list = l
    s_length = begin_length
    while s_length <= max_length: #保证每个组合的长度不超过最大长度值
        s_pos= 0 #每一个元素的位置
        while s_pos < max_length: #处理每一个元素,s_pos为当前元素的位置
            allSegment = getAllSegment(s_length,data_list,s_pos,1)
            if len(allSegment)>0:
                print(allSegment)
            s_pos = s_pos + 1
        s_length = s_length + 1 #组合的长度增加

#获取所有的elem对应的数组，并且组合
def getAllElemAllSegmentNoRepeat(begin_length,max_length,l):
    data_list = list(l)
    s_length = begin_length
    while s_length <= max_length: #保证每个组合的长度不超过最大长度值
        s_pos= 0 #每一个元素的位置
        while s_pos < len(data_list): #处理每一个元素,s_pos为当前元素的位置
            allSegment = getAllSegment(s_length,data_list,s_pos,1)
            if len(allSegment)>0:
                print(allSegment)
            del (data_list[0])
        s_length = s_length + 1 #组合的长度增加
        data_list = list(l)

l=[1,2,3,4,5,6,7,8,9]
max_length = len(l)
s_length = 0 #组合的长度变量

# getAllElemAllSegmentNoRepeat(0,len(l),l)
logic([1,2],2.5,1.5)
logic([2,3],2.5,1.5)
logic([3,1],2.5,1.5)


