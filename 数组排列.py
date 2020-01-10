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

#��ȡstart_pos��ʼ�㵽end_pos�յ������
def getSegment(start_pos,end_pos,l):
    segment = list()
    #�ж��Ƿ񳬳���Χ
    if end_pos >= len(l):
        return segment

    # ��ʼ�Դ�ÿ�����
    local_start_pos = start_pos
    local_end_pos = end_pos

    while local_start_pos <= local_end_pos:
        segment.append(l[local_start_pos])
        local_start_pos = local_start_pos + 1

    # �����Դ�ÿ�����
    return segment

#��ȡ���е�ĳ���ȵ�����
def getAllSegment(s_length,l,s_pos,with_value):
    local_start_pos = 0
    local_end_pos = s_length-1

    allSegment = list()

    #���⴦��
    if s_length == 0:
         simpleList = [l[s_pos]]
         allSegment.append(simpleList)
         return allSegment

    # ͨ��localEnd��ֵ�ж��Ƿ������ ���� ȥ������Ҫ����ֵ
    while local_end_pos < len(l):
        if (local_start_pos > s_pos and local_end_pos > s_pos or local_start_pos < s_pos and local_end_pos < s_pos):
            segment =getSegment(local_start_pos,local_end_pos,l)
            if len(segment) > 0:
                #���with_valueΪ1���Ͱ���s_pos����ֵ
                if with_value == 1:
                    segment.append(l[s_pos])
                allSegment.append(segment)

        local_end_pos = local_end_pos + 1
        local_start_pos = local_start_pos + 1

    return allSegment

#��ȡ���е�elem��Ӧ�����飬�������
def getAllElemAllSegment(begin_length,max_length,l):
    data_list = l
    s_length = begin_length
    while s_length <= max_length: #��֤ÿ����ϵĳ��Ȳ�������󳤶�ֵ
        s_pos= 0 #ÿһ��Ԫ�ص�λ��
        while s_pos < max_length: #����ÿһ��Ԫ��,s_posΪ��ǰԪ�ص�λ��
            allSegment = getAllSegment(s_length,data_list,s_pos,1)
            if len(allSegment)>0:
                print(allSegment)
            s_pos = s_pos + 1
        s_length = s_length + 1 #��ϵĳ�������

#��ȡ���е�elem��Ӧ�����飬�������
def getAllElemAllSegmentNoRepeat(begin_length,max_length,l):
    data_list = list(l)
    s_length = begin_length
    while s_length <= max_length: #��֤ÿ����ϵĳ��Ȳ�������󳤶�ֵ
        s_pos= 0 #ÿһ��Ԫ�ص�λ��
        while s_pos < len(data_list): #����ÿһ��Ԫ��,s_posΪ��ǰԪ�ص�λ��
            allSegment = getAllSegment(s_length,data_list,s_pos,1)
            if len(allSegment)>0:
                print(allSegment)
            del (data_list[0])
        s_length = s_length + 1 #��ϵĳ�������
        data_list = list(l)

l=[1,2,3,4,5,6,7,8,9]
max_length = len(l)
s_length = 0 #��ϵĳ��ȱ���

# getAllElemAllSegmentNoRepeat(0,len(l),l)
logic([1,2],2.5,1.5)
logic([2,3],2.5,1.5)
logic([3,1],2.5,1.5)


