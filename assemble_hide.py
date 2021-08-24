import win32com.client as win32

def assemble_hide():  # 隱藏組力拘束
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    # ===============================搜尋拘束==============================
    selection1 = document.Selection
    selection1.Search("Name=Constraints,all")
    visPropertySet1 = selection1.VisProperties
    visPropertySet1 = visPropertySet1.Parent
    bSTR1 = visPropertySet1.Name
    visPropertySet1.SetShow(1)
    selection1.Clear()
    # ===============================搜尋拘束==============================
    # ===============================搜尋組立起點==============================
    selection1.Search("Name=bned_up_forming_boit_point_*,all")
    visPropertySet2 = selection1.VisProperties
    bSTR2 = visPropertySet2.Name
    visPropertySet2.SetShow(1)
    selection1.Clear()
    # ===============================搜尋組立起點==============================
