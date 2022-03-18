trgt='E:\\GIT_REP\\automation-python-scripts\\reports\\win32\\test\\Еженедельный отчет 04-10_08-10-2021.docx'



def write_to_file(target):
    import win32com.client as win32
    word = win32.gencache.EnsureDispatch('Word.Application')
    my_doc1=word.Documents.Open(target)
    #my_doc1.Visible = 1
    #my_doc1.Visible = 1
    tables = my_doc1.Tables.Count
    print(tables)
    for table in range(1,tables+1):
        try:
            print(str(my_doc1.Tables(table).Cell(1,1).Range.Text))
        except:
            print("close")
        if str('Группа') in str(my_doc1.Tables(table).Cell(1,1).Range.Text):
            my_doc1.Tables(table).Cell(2,2).Range.Text = 'test'
            my_doc1.Tables(table).Cell(3,2).Range.Text = 'test'
            my_doc1.Tables(table).Cell(4,2).Range.Text = 'test'
            
            #my_doc1.Quit()
            print("done")
        else:
            print("error")
    my_doc1.Close()
write_to_file(trgt)
