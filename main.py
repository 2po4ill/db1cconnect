# coding=cp1251

import pythoncom
import win32com.client

V83_CONN_STRING = 'Srvr="zck1c8devsrv2:2541";Ref="10703";";Usr="aa206";Pwd="";'
pythoncom.CoInitialize()
V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
q1 = '''
ВЫБРАТЬ
  Док.Дата КАК TN_DATE,
  Док.Номер КАК TN_NUMBER,
  Док.Подразделение.Наименование КАК TN_UNIT_TITLE,
  Док.Склад.Наименование КАК TN_STORE_TITLE,
  Док.ДокументОснование.Номер КАК ZNP_NUMBER,
  Док.ДокументОснование.Дата КАК ZNP_DATE
ИЗ
  Документ.ТребованиеНакладная КАК Док
ГДЕ
  Док.Ссылка.Дата = ДАТА(2019, 2, 25)  #Вот тут нужно сравнивать Док.Ссылка.Дата с текущей датой, а не писать каждый раз ручками!
'''