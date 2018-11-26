import xlrd,pymysql,time
from xlutils.copy import copy
import traceback
## ##读取目录下文件的第一页
##确文件的第六列跟第七列分别是离职与交接人的email
## ##
TIMEFORMAT="%Y-%m-%d %X"
data =xlrd.open_workbook('YW05-离职流程报表 (32).xls')
table = data.sheets()[0] ##通过索引顺序获取
##table = data.sheet_by_index(0) #通过索引顺序获取
##table =data.sheet_by_name(u'Sheet1')#通过名称获取 ##  获取整行和整列的值（数组）

try:
    ##测试环境and '@' in table.row_values(line)[9]
##        conn_online=pymysql.connect(host='****',user='****',passwd='****',db='****',port=3306,charset='utf8')
    ##线上环境
        conn_online=pymysql.connect(host='rdsc14z37q7i8pq2m79fo.mysql.rds.aliyuncs.com',user='****',passwd='****',db='****',port=3306,charset='utf8')      
        cur_online=conn_online.cursor()#获取一个游标
##    cur_test=conn_test.cursor()
except  Exception as  e:
          print("发生异常",e)

for line in range(1,table.nrows):
    if table.row_values(line)[4] :
        quit_name =  table.row_values(line)[3]
        quit_email =  table.row_values(line)[5]        
        jie_name = table.row_values(line)[6]
        jie_email = table.row_values(line)[7]        
        
        try:
            sql_check_quit='select * from t_customer_aid_quit where quit_customer_email= "' +quit_email+'"'
            cur_online.execute(sql_check_quit)
            data=cur_online.fetchall()
            if len(data) != 0 and data:###处理已经离职过的数据
##                table.row_values(line)[13]='离职账号'+quit_email+'已经离过职';
##                wb = copy(table)
##                ws = wb.get_sheet(0)
##                ws.write(line,13,'离职账号'+quit_email+'  已经存在')
                sql_check_quit_2='select * from t_customer_aid where email= "' +jie_email+'"'
                cur_online.execute(sql_check_quit_2)
                data_2=cur_online.fetchall()
                if  len(data_2)==0:
                        print('接收人'+jie_email+' 不存在，未做离职处理')
                        continue
                if data_2 and data_2[0][7] ==3:
                    print('交接人'+data_2[0][2]+data_2[0][1]+'是分总，未做离职处理   ')
                    continue
                if data_2  and data_2[0][1] != jie_name:
                        if data_2[0][1] != jie_name+'1':
                            print('接收人：'+data_2[0][2]+':'+data_2[0][1]+'与 excel ：'+jie_email+':'+jie_name+' 姓名不一致!未做离职处理')
                            continue
                print("UPDATE t_customer_aid_quit SET  take_customer_id='"+str(data_2[0][0])+"', take_customer_name='"+data_2[0][1]+"', take_customer_email='"+data_2[0][2]+"',quit_time='"+time.strftime( TIMEFORMAT, time.localtime( ) )+"' ,account_process_status='0', relation_process_status='1', task_process_status='0' WHERE id='"+str(data[0][0])+"';") 
                continue
##            sql_onlie1 ='select id,name,email from t_customer_aid where email="'+jie_email+'"'
##            status = cur_online.execute(sql_online1)
##            data=cur_online.fetchall()
##            if len(data) == 0:
##                 print('接收人账号'+jie_email +' 不存在')
##                 continue
###未离职处理
            sql_online ='select a.id,a.name,a.email,b.id,b.name,b.email,"0", "1", "0", "1",a.role from t_customer_aid a,(select id,name,email from t_customer_aid where email="'+jie_email+'") b  where a.email="'+quit_email+'"'
            status = cur_online.execute(sql_online)
            ##          print('sql_online: ')
            ##          print(sql_test)
            data=cur_online.fetchall()
            if len(data) == 0:
                print('接收人'+jie_email+' 或关闭助手人'+quit_email+' 不存在，未做离职处理')
                continue         
            for d in data :
                if d[1] != quit_name:
                    print('关闭助手权限人：'+d[2]+':'+d[1]+'与 excel ：'+quit_email+':'+quit_name+' 姓名不一致')
                    continue
                if d[4] != jie_name:
                    print('接收人：'+d[4]+':'+d[5]+'与 excel ：'+jie_email+':'+jie_name+' 姓名不一致')
                    continue
##                if d[10] == 3:
##                        sql_online = ''
##                        if 
##                    print('交接人'+d[4]+':'+d[5]+'是分总，未做离职处理   ')
##                    continue
                sql_test='insert into t_customer_aid_quit(quit_customer_id,quit_customer_name,quit_customer_email,take_customer_id,take_customer_name,take_customer_email,quit_time,account_process_status,relation_process_status,task_process_status,create_time,state) '
                sql_test +='values("'+str(d[0])+'","'+str(d[1])+'","'+str(d[2])+'","'+str(d[3])+'","'+str(d[4])+'","'+str(d[5])+'","'+\
                time.strftime( TIMEFORMAT, time.localtime( ) )+'","'+str(d[6])+'","'+str(d[7])+'","'+str(d[8])+'","'+time.strftime( TIMEFORMAT, time.localtime( ) )+'","'+str(d[9])+'")'
                ##              cur_test.execute(sql_test)
                ##              conn_test.commit()
                ##              print('sql_test: ')
                print(sql_test+';')
        except  Exception as  e:
            print("发生异常",e)
            traceback.print_exc()
    else:
        print("email没有@")

cur_online .close()#关闭游标
##cur_test.close()
conn_online.close()#释放数据库资源
##conn_test.close()          
print("总计数量："+str(line))
    
##print(table.col_values(1))
