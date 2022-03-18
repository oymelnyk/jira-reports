from docx import Document
import os
from jira import JIRA
from datetime import datetime, date, timedelta
from pathlib import Path
import shutil
import time
from getpass import getpass
import win32com.client as win32


Save_Path = "P:\\2022\\" #Сетевой диск
months = {"01":"январь", "02":"февраль", "03":"март", "04":"апрель", "05":"май", "06":"июнь", "07":"июль", "08":"август", "09":"сентябрь", "10":"октябрь", "11":"ноябрь", "12":"декабрь"} # , "13":"month"} 
print("login:")
login = input() #LOGIN
pwd = getpass()

def current_month(Save_Path, months):
    
    today = date.today().strftime("%m")
    print(today)
    print(months[str(today)])
    #today = "08"
    current_save_path = Save_Path+months[str(today)] #Save_Path+months["13"]


    def copy_file(current_save_path, target_save_path):
        mass = []

        old_files_names = os.listdir(path=current_save_path)
        for i in old_files_names:
            gtime = os.path.getmtime(current_save_path+"\\\\"+i)
            atime = datetime.fromtimestamp(gtime)
            mass.append(str(atime)+" "+str(i))
        mass.sort()
        mass_split = mass[-1].split()
        
        old_file_name = (mass_split[2]+" "+mass_split[3]+" "+mass_split[4])
        print(old_file_name)
        original = current_save_path+"\\\\"+old_file_name
        get_now = datetime.now()
        plus_friday = get_now + timedelta(days=4)
        minus_5 = plus_friday - timedelta(days=4) #отнимаем дни до понедельника
        minus_5_day = datetime.strftime(minus_5, "%d-%m")
        date_now = datetime.strftime(get_now, "%d-%m-%Y")
        plus_day_to_friday = datetime.strftime(plus_friday, "%d-%m-%Y")
        print(minus_5_day)
        print(date_now)
        print(plus_day_to_friday)
        global target
        target = target_save_path+"\\\\Еженедельный отчет "+str(minus_5_day)+"_"+str(plus_day_to_friday)+".docx" 

        shutil.copyfile(original, target) # копируем прошлый файл
        return target

    try:
        
        os.mkdir(current_save_path)
        print("Создана директория: "+current_save_path)
        return current_save_path
        
    except:
        print("Директория "+current_save_path+" уже существует.")

        # Если в директории по заданому пути нет файлов, значит ищем файл прошлого месяца
    if not os.listdir(path=current_save_path):
        print("нет файла")
        low_date = int(today)-1
        if low_date<10:
            low_date = "0"+str(low_date)
            
        elif low_date<0:
            low_date = "12"
        
        copy_file(Save_Path+months[str(low_date)], current_save_path)



        # Если в директории по заданому пути есть файлы, значит ищем самый новый файл и копируем его.
    else:
        copy_file(current_save_path, current_save_path)

    return target
        
        
from getpass import getpass
def get_jira_info(login, pwd):
    #print("login: ")
    
    jira_server = 'https://helpdesk.sbrf.com.ua/jira'
    jira= JIRA(server=jira_server, basic_auth=(login, pwd))

    assigned_groups = ['jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line', 'jira_it_GroupName_2line']
    
    global done_issues
    global in_progress_issues
    global canceled_issues
    global customer_approval


    done_issues = []
    in_progress_issues = []
    canceled_issues = []
    customer_approval = []


    #jql_done = 'project = ITSD AND status in (Resolved, "Customer Approval") AND updated >= -7d AND updated <= 0d AND "Assigned group" = '
    #jql_in_progress = 'project = ITSD AND status in (Open, Reopened, "Waiting for customer", "Waiting for approval", "Work in progress", "Work in outsource", "Open 2 line") AND updated >= -7d AND updated <= 0d AND "Assigned group" = '
    #jql_canceled = 'project = ITSD AND status in (Closed, Canceled) AND updated >= -7d AND updated <= 0d AND "Assigned group" = '

    for assigned_group in assigned_groups:
        jql_done = jira.search_issues('project = ITSD AND status in (Resolved) AND updated >= -7d AND updated <= 0d AND "Assigned group" = '+ str(assigned_group), maxResults=10000)
        jql_in_progress = jira.search_issues('project = ITSD AND status in (Open, Reopened, "Waiting for customer", "Waiting for approval", "Work in progress", "Work in outsource", "Open 2 line") AND "Assigned group" = '+ str(assigned_group), maxResults=10000)
        jql_canceled = jira.search_issues('project = ITSD AND status in (Closed, Canceled) AND updated >= -7d AND updated <= 0d AND "Assigned group" = '+ str(assigned_group), maxResults=10000)
        jql_customer_approval = jira.search_issues('project = ITSD AND status in ("Customer Approval") AND updated >= -7d AND updated <= 0d AND "Assigned group" = '+ str(assigned_group), maxResults=10000)
        done_issues.append(len(jql_done))
        in_progress_issues.append(len(jql_in_progress))
        canceled_issues.append(len(jql_canceled))
        customer_approval.append(len(jql_customer_approval))

    print("done")
    print(done_issues)
    print(in_progress_issues)
    print(canceled_issues)
    print(customer_approval)
    return done_issues, in_progress_issues, canceled_issues, customer_approval


#def write_to_file(target, done_issues, in_progress_issues, canceled_issues):
#    document = Document(target)
#    table_mass = []
#    for table in document.tables:
#        hdr_cells = table.rows[0].cells
#        if hdr_cells[0].text == 'Группа':
#            table_mass.append(table)
#
#
#   for i in range(1,10):
#
#        hdr_cells = table_mass[0].rows[i].cells
#        ii = i-1
#        hdr_cells[1].text = str(done_issues[ii])
#        hdr_cells[2].text = str(in_progress_issues[ii])
#        hdr_cells[3].text = str(canceled_issues[ii])
#
#    document.save(target)
#    document.save(target)
#    document.save(target)
#    print(target)

def write_to_file(target, done_issues, in_progress_issues, canceled_issues, customer_approval):
    
    word = win32.gencache.EnsureDispatch('Word.Application')
    my_doc1=word.Documents.Open(target)
    my_doc1.Visible = 0
    #my_doc1.Visible = 1
    tables = my_doc1.Tables.Count
    print(tables)
    for table in range(1,tables+1):
        if str('Группа') in str(my_doc1.Tables(table).Cell(1,1).Range.Text):
            for i in range(2,11):
                ii = i-2
                my_doc1.Tables(table).Cell(i,2).Range.Text = str(done_issues[ii]) #1 ряд, 2 столбик
                my_doc1.Tables(table).Cell(i,3).Range.Text = str(in_progress_issues[ii])
                my_doc1.Tables(table).Cell(i,4).Range.Text = str(canceled_issues[ii])
                my_doc1.Tables(table).Cell(i,5).Range.Text = str(customer_approval[ii])
            
            #my_doc1.Quit()
            print("done")
        else:
            print("error")
    my_doc1.Close()





while True:
    dtn = datetime.now()
    check_day = int(datetime.strftime(dtn, "%w"))
    check_time = int(datetime.strftime(dtn, "%H"))
    #check_day
    if check_day == 1 and check_time == 17: #and check_time == 11:
       try:
           current_month(Save_Path, months)
           
       except:
          print("error")
    else:
        print('Время не наступило или - file is exist')
    if check_day == 5 and check_time == 15:
        mass=[]
        today = date.today().strftime("%m")
        current_save_path = Save_Path+months[str(today)]
        print(current_save_path)
        old_files_names = os.listdir(path=current_save_path)
        for i in old_files_names:
            gtime = os.path.getmtime(current_save_path+"\\\\"+i)
            atime = datetime.fromtimestamp(gtime)
            mass.append(str(atime)+" "+str(i))
        mass.sort()
        mass_split = mass[-1].split()
        
        old_file_name = (mass_split[2]+" "+mass_split[3]+" "+mass_split[4])
        print(old_file_name)
        original = current_save_path+"\\\\"+old_file_name
        get_now = datetime.now()
        minus_5 = get_now - timedelta(days=4) #отнимаем дни до понедельника
        minus_5_day = datetime.strftime(minus_5, "%d-%m")
        date_now = datetime.strftime(get_now, "%d-%m-%Y")
        print(minus_5_day)
        print(date_now)
        global target
        target = current_save_path+"\\\\Еженедельный отчет "+str(minus_5_day)+"_"+str(date_now)+".docx" 




        get_jira_info(login, pwd)
        write_to_file(target,done_issues, in_progress_issues, canceled_issues, customer_approval)
    else:
        ('пропуск')
    print("time sleep - "+str(dtn))
    time.sleep(3600)







