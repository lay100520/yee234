import win32com.client as win32
catapp = win32.Dispatch('CATIA.Application')
documents1 = catapp.Documents
try :
    for i in range(99):
        partDocumnet = catapp.ActiveDocument
        partDocumnet.Close()
except:
    pass