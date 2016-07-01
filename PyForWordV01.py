#-*- coding: UTF-8 -*-
import win32com
from win32com.client import Dispatch, constants
w = win32com.client.Dispatch('Word.Application')
w.Visible = 0
w.DisplayAlerts = 0
doc = w.Documents.Open( FileName = "D:ShubaoLv-DOC/test.doc" )
doc.Tables[0].Rows[0].Cells[0].Range.Text = u'东北林业大学'
doc.Tables[1].Rows[1].Cells[1].Range.Text = u'acm亚洲赛'
doc.Tables[2].Rows[2].Cells[2].Range.Text = u'ACM-ICPC'
doc.Tables[0].Rows.Add() # 增加一行
doc.Close()
w.Quit()
