# coding=cp1251

import pythoncom
import win32com.client

V83_CONN_STRING = 'Srvr="zck1c8devsrv2:2541";Ref="10703";";Usr="aa206";Pwd="";'
pythoncom.CoInitialize()
V83 = win32com.client.Dispatch("V83.COMConnector").Connect(V83_CONN_STRING)
q1 = '''
�������
  ���.���� ��� TN_DATE,
  ���.����� ��� TN_NUMBER,
  ���.�������������.������������ ��� TN_UNIT_TITLE,
  ���.�����.������������ ��� TN_STORE_TITLE,
  ���.�����������������.����� ��� ZNP_NUMBER,
  ���.�����������������.���� ��� ZNP_DATE
��
  ��������.������������������� ��� ���
���
  ���.������.���� = ����(2019, 2, 25)  #��� ��� ����� ���������� ���.������.���� � ������� �����, � �� ������ ������ ��� �������!
'''