import pythoncom
pythoncom.CoInitialize()
import clr
import  mouse
from time import sleep
import wx
import wx.stc as stc
import os
import subprocess as sp
import threading
import unicodedata
import webbrowser
import pythoncom

class EditorFrame(wx.Frame): #wxPythonの設定
    def __init__(self):
        #ハイライトするキーワード
        highlight_keyword1 = ["import","from","class","def","while","for","if","elif","else"\
            ,"break","continue","pass","try","except","with","as","is","__name__","in","global","return"\
                ,"True","False","None","and","or","not"]

        highlight_keyword2 = ["abs","all","any","ascii","bin","bool","breakpoint","bytearray","bytes","callable","chr","classmethod","compile","complex"\
            ,"delattr","dict","dir","divmod","enumerate","eval","exec","filter","float","format","frozenset","getattr","globals","hasattr"\
                ,"hash","help","hex","id","input","int","isinstance","issubclass","iter","len","list","locals","map","max"\
                    ,"memoryview","min","next","object","oct","open","ord","pow","print","property","range","repr","reversed","round"\
                        ,"set","setattr","slice","sorted","staticmethod","str","sum","super","tuple","type","vars","zip","__import__"]

        wx.Frame.__init__(self,None,-1,"EyEs Editor",size=(1000,1000))
        self.panel = wx.Panel(self,wx.ID_ANY)

        self.text = stc.StyledTextCtrl(self.panel, -1,
                                     style=wx.TE_MULTILINE)

        self.text.SetLexer(stc.STC_LEX_PYTHON)
        self.text.SetThemeEnabled(True)
        self.text.StyleSetSpec(stc.STC_STYLE_DEFAULT,"size:30,face:UD デジタル 教科書体 NP-R")
        self.text.StyleClearAll()
        
        self.text.SetCaretStyle(1)
        self.text.SetCaretWidth(10)
        self.text.SetCaretPeriod(100)
        self.text.SetCaretSticky(100)
        self.text.SetEdgeColumn(10)

        self.text.StyleSetForeground(stc.STC_P_IDENTIFIER,wx.Colour("#FFFFFF"))
        self.text.StyleSetBackground(stc.STC_P_IDENTIFIER,wx.Colour("#0072B2"))

        faces = {
                  "font" : "face:UD デジタル 教科書体 NP-R",
                  "size" : 30,
                  }

        fonts = "face:%(font)s,size:%(size)d" % faces
        self.text.StyleSetSpec(stc.STC_STYLE_DEFAULT,"size:30,face:UD デジタル 教科書体 NP-R")
        self.text.StyleSetBackground(stc.STC_STYLE_DEFAULT,wx.Colour("#FFFFFF"))
        self.text.StyleSetBackground(stc.STC_P_DEFAULT,wx.Colour("#FFFFFF"))
        self.text.StyleSetForeground(stc.STC_P_DEFAULT,wx.Colour("#000000"))
        self.text.SetKeyWords(0," ".join(highlight_keyword1))
        self.text.SetKeyWords(1," ".join(highlight_keyword2))
        self.text.StyleSetForeground(stc.STC_P_WORD, wx.Colour("#000000"));
        self.text.StyleSetForeground(stc.STC_P_WORD2, wx.Colour("#0072B2"));
        self.text.StyleSetBackground(stc.STC_P_WORD, wx.Colour("#FFA500"));
        self.text.StyleSetBackground(stc.STC_P_WORD2, wx.Colour("#FFA500"));

        self.text.StyleSetSpec(stc.STC_P_COMMENTLINE,"fore:#CC79A7,back:#FFFFFF" + fonts)
        self.text.StyleSetSpec(stc.STC_PAS_COMMENT, "fore:#CC79A7,back:#FFFFFF" + fonts)
        self.text.StyleSetSpec(stc.STC_P_STRING, "fore:#000000,back:#0072B2" + fonts)
        self.text.StyleSetSpec(stc.STC_P_CHARACTER, "fore:#000000,back:#0072B2" + fonts)
        self.text.StyleSetSpec(stc.STC_P_STRINGEOL,"fore:#000000,back:#0072B2" + fonts)
        self.text.StyleSetSpec(stc.STC_P_COMMENTBLOCK,"fore:#CC79A7,back:#FFFFFF" + fonts)
        self.text.StyleSetSpec(stc.STC_P_TRIPLEDOUBLE,"fore:#CC79A7,back:#FFFFFF" + fonts)
        self.text.StyleSetSpec(stc.STC_P_DEFNAME,"fore:#E69F00,back:#0072B2" + fonts)
        self.text.StyleSetSpec(stc.STC_P_CLASSNAME,"fore:#E69F00,back:#0072B2" + fonts)
        self.text.StyleSetSpec(stc.STC_P_NUMBER, "fore:#56B4E9,back:#0072B2" + fonts)
        self.text.StyleSetSpec(stc.STC_P_OPERATOR, "fore:#FFA500,back:#0072B2" + fonts)
        self.text.SetMarginType(3, stc.STC_MARGIN_NUMBER)
        self.text.SetMarginWidth(3, 60)
        self.text.StyleSetSpec(stc.STC_STYLE_LINENUMBER,  "fore:#000000")
        self.text.SetIndent(4)
        self.text.SetMarginBackground(0,wx.Colour("#FFFFFF"))

        self.eyes_text = wx.TextCtrl(self.panel,-1,"",style=wx.TE_MULTILINE)
        font_eyes_text = wx.Font(30, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,faceName="UD デジタル 教科書体 NP-R")
        self.eyes_text.SetFont(font_eyes_text)

        self.button_run = wx.Button(self.panel,-1,"実行")
        self.button_1 = wx.Button(self.panel,-1,"戻る")
        self.button_2 = wx.Button(self.panel,-1,"進む")
        self.button_3 = wx.Button(self.panel,-1,"新規\n作成")
        self.button_4 = wx.Button(self.panel,-1,"開く")
        self.button_5 = wx.Button(self.panel,-1,"保存")
        self.button_6 = wx.Button(self.panel,-1,"名前\n保存")
        self.button_7 = wx.Button(self.panel,-1,"拡大")
        self.button_8 = wx.Button(self.panel,-1,"縮小")

        font_button = wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,faceName="UD デジタル 教科書体 NP-R")

        self.button_run.SetFont(font_button)
        self.button_1.SetFont(font_button)
        self.button_2.SetFont(font_button)
        self.button_3.SetFont(font_button)
        self.button_4.SetFont(font_button)
        self.button_5.SetFont(font_button)
        self.button_6.SetFont(font_button)
        self.button_7.SetFont(font_button)
        self.button_8.SetFont(font_button)

        self.button_run.Bind(wx.EVT_BUTTON,self.python_run)
        self.button_1.Bind(wx.EVT_BUTTON,self.text_undo)
        self.button_2.Bind(wx.EVT_BUTTON,self.text_redo)
        self.button_3.Bind(wx.EVT_BUTTON,self.new_text)
        self.button_4.Bind(wx.EVT_BUTTON,self.open_text)
        self.button_5.Bind(wx.EVT_BUTTON,self.save_text)
        self.button_6.Bind(wx.EVT_BUTTON,self.name_save_text)
        self.button_7.Bind(wx.EVT_BUTTON,self.text_zoom1)
        self.button_8.Bind(wx.EVT_BUTTON,self.text_zoom2)


 
        self.SetScrollbar(wx.VERTICAL, 0, 16, 16)
        self.SetScrollbar(wx.HORIZONTAL, 0, 16, 16)

        self.text_zoom_level = 0

        self.main_sizer = wx.FlexGridSizer(2,2,10,0)
        self.sub_sizer_1 = wx.GridSizer(8,1,10,0)
        self.main_sizer.AddGrowableCol(1)
        self.main_sizer.AddGrowableRow(1)

        self.sub_sizer_1.Add(self.button_1,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_2,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_3,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_4,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_5,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_6,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_7,0, flag=wx.GROW)
        self.sub_sizer_1.Add(self.button_8,0, flag=wx.GROW)

        self.main_sizer.Add(self.button_run,0, flag=wx.GROW)
        self.main_sizer.Add(self.eyes_text,0, flag=wx.EXPAND)
        self.main_sizer.Add(self.sub_sizer_1,0,flag=wx.GROW)
        self.main_sizer.Add(self.text,0, flag=wx.EXPAND)

        self.panel.SetSizer(self.main_sizer)
        self.menu_file = wx.Menu()
        self.menu_file.Append(1, "新規作成")
        self.menu_file.Append(2, "開く")
        self.menu_file.Append(3, "保存")
        self.menu_file.Append(4, "名前を付けて保存")
        self.menu_file.Append(5, "終了")

        self.menu_edit = wx.Menu()
        self.menu_edit.Append(6,"戻る")
        self.menu_edit.Append(7,"進む")
        self.menu_edit.Append(8,"切り取り")
        self.menu_edit.Append(9,"コピー")
        self.menu_edit.Append(10,"貼り付け")

        self.menu_run = wx.Menu()
        self.menu_run.Append(11,"実行")

        self.menu_option = wx.Menu()
        self.menu_option.Append(13,"拡大")
        self.menu_option.Append(14,"縮小")

        self.menu_help = wx.Menu()
        self.menu_help.Append(18,"ダウンロードページを開く")

        self.menu_bar = wx.MenuBar()
        self.menu_bar.Append(self.menu_file, "ファイル")
        self.menu_bar.Append(self.menu_edit, "編集")
        self.menu_bar.Append(self.menu_run, "実行")
        self.menu_bar.Append(self.menu_option, "設定")
        self.menu_bar.Append(self.menu_help,"ヘルプ")

        self.SetMenuBar(self.menu_bar)
        self.Bind(wx.EVT_MENU, self.click_menu)

        self.id = wx.NewIdRef()
        self.RegisterHotKey(self.id,wx.MOD_SHIFT,virtualKeyCode=116)
        self.Bind(wx.EVT_HOTKEY,self.python_run,id=self.id)

        self.id = wx.NewIdRef()
        self.RegisterHotKey(self.id,wx.MOD_CONTROL,virtualKeyCode=78)
        self.Bind(wx.EVT_HOTKEY,self.new_text,id=self.id)

        self.id = wx.NewIdRef()
        self.RegisterHotKey(self.id,wx.MOD_CONTROL,virtualKeyCode=83)
        self.Bind(wx.EVT_HOTKEY,self.save_text,id=self.id)

        self.id = wx.NewIdRef()
        self.RegisterHotKey(self.id,wx.MOD_CONTROL,virtualKeyCode=83)
        self.Bind(wx.EVT_HOTKEY,self.name_save_text,id=self.id)

        self.Bind(wx.EVT_CLOSE,self.exit_window)

        self.pre_save_text()
        self.Show()

    def cursor_text(self):
        return str(self.text.GetCurLine()[0])

    def click_menu(self,event):
        event_id = event.GetId()
        if event_id == 1:
            self.new_text(0)
        elif event_id == 2:
            self.open_text(0)

        elif event_id == 3:
            self.save_text(0)

        elif event_id == 4:
            self.name_save_text(0)

        elif event_id == 5:
            self.exit_window(0)

        elif event_id == 6:
            self.text.Undo()
        
        elif event_id == 7:
            self.text.Redo()
        
        elif event_id == 8:
            self.text.Cut()

        elif event_id == 9:
            self.text.Copy()

        elif event_id == 10:
            self.text.Paste()

        elif event_id == 11:
            self.python_run(0)

        elif event_id == 13:
            self.text_zoom1(0)       
                
        elif event_id == 14:
            self.text_zoom2(0)    

        elif event_id == 18:
            webbrowser.open("https://github.com/yoko1004/EyEs_Editor/releases")

    def pre_save_text(self):
        with open("config/pre_save.txt","r",encoding="utf-8") as p:
            pre_save_path = p.read()
            if pre_save_path == "n":
                self.text.SetValue("ようこそ")
                with open("config/pre_save.txt", "w", encoding="utf-8") as f:
                    f.write("")
            elif pre_save_path != "":
                with open(pre_save_path,"r",encoding="utf-8") as p2:
                    self.text.SetValue(p2.read())

    def new_text(self,event):
        with open("config/pre_save.txt","w",encoding="utf-8") as n:
            n.write("")
            self.text.SetValue("")
            self.SetTitle("EyEs Editor --- 未保存")

    def open_text(self,event):
        filter = "python file(*.py;*.pyw) | *.py;*.pyw | All file(*.*) | *.*"
        open_dlg = wx.FileDialog(self, u"開く","os.getcwd()","",filter,style=wx.FD_OPEN)
        if open_dlg.ShowModal() == wx.ID_OK:
            self.filename = open_dlg.GetFilename()
            self.dirname = open_dlg.GetDirectory()
            with open(os.path.join(self.dirname, self.filename),"r",encoding="utf-8") as o:
                self.text.SetValue(o.read())
                self.SetTitle("EyEs Editor --- " + str(self.filename))
            with open("config/pre_save.txt","w",encoding="utf-8") as s:
                s.write(self.dirname + "\\" + self.filename)
        open_dlg.Destroy()

    def save_text(self,event):
        with open("config/pre_save.txt","r",encoding="utf-8") as s:
            save_path = s.read()
            if save_path != "":
                with open(save_path, "w", encoding="utf-8") as f:
                    print("OK")
                    f.write(self.text.GetValue())
            else:
                self.name_save_text(0)

    def name_save_text(self,event):

        filter = "python file(*.py;*.pyw) | *.py;*.pyw | All file(*.*) | *.*"
        name_save_dlg = wx.FileDialog(self, u"名前を付けて保存","os.getcwd()","",filter,style=wx.FD_SAVE)
        if name_save_dlg.ShowModal() == wx.ID_OK:
            self.filename = name_save_dlg.GetFilename()
            self.dirname = name_save_dlg.GetDirectory()
            with open(os.path.join(self.dirname, self.filename), "w", encoding="utf-8") as f:
                f.write(self.text.GetValue())
            with open("config/pre_save.txt","w",encoding="utf-8") as s:
                s.write(self.dirname + "\\" + self.filename)
        name_save_dlg.Destroy()

    def python_run(self,event):
        self.save_text(0)
        with open("config/pre_save.txt","r",encoding="utf-8") as p:
            now_path = p.read()
            if now_path != "":
                sp.Popen(["start",os.path.abspath("config/python_run.bat"), now_path],universal_newlines=True,shell=True)
            else:
                pass
    
    def text_zoom1(self,event):
        self.text_zoom_level += 2
        self.text.SetZoom(self.text_zoom_level)

    def text_zoom2(self,event):
        self.text_zoom_level -= 2
        self.text.SetZoom(self.text_zoom_level)

    def text_undo(self,event):
        self.text.Undo()

    def text_redo(self,event):
        self.text.Redo()

    def exit_window(self,event):
        exit_dlg = wx.MessageDialog(self, "プログラムを終了しますか？","修了確認",wx.YES_NO)
        if exit_dlg.ShowModal() == wx.ID_YES:
            pythoncom.CoUninitialize()
            wx.Exit()
            exit()

def text_count(texts): #自動改行
    c = 0
    text_list = list()
    for w in texts:
        if unicodedata.east_asian_width(w) in "FWA":
            c += 2
        else:
            c += 1
        text_list.append(w)
        if c >= 42:
            text_list.append("\n")
            c = 0
    texts = "".join(text_list)
    return texts

def get_text(): #テキスト取得
    global editor_frame
    path = str(os.path.abspath("config/") +"\\getText")
    clr.AddReference(path)
    from Gettext import Gettext1 #DLL呼び出し
    gettext1 = Gettext1()
    old_pos = ()
    old_ui = ""
    old_msaa = ""
    old_cursor_text = ""
    count = 0
    while True:
        new_pos = mouse.get_position() #マウスの画面座標を取得

        if new_pos != old_pos:
            old_pos = new_pos

            new_ui = str(gettext1.GetElementFromCursorByUI(old_pos[0],old_pos[1])) #UIAUTOMATIONを利用してテキストを入手
            new_msaa = str(gettext1.GetElementFromCursorByMSAA(old_pos[0],old_pos[1])) #Microsoft Active Accessibilityを利用してテキストを入手

            #UIAUTOMATIONの利用が推奨されているので優先度が高い
            if new_ui != "" and new_ui != None and new_ui != "None" and new_ui != old_ui: #空白チェック
                old_ui = new_ui
                editor_frame.eyes_text.SetLabel(text_count(old_ui)) #描写
                count = 0

            elif new_msaa != "" and new_msaa != None and new_msaa != "None" and new_msaa != old_msaa: #空白チェック
                old_msaa = new_msaa
                editor_frame.eyes_text.SetLabel(text_count(old_msaa)) #描写
                count = 0

        new_cursor_text = editor_frame.cursor_text() #エディタの文を取得
        if new_cursor_text != old_cursor_text:
            old_cursor_text = new_cursor_text
            editor_frame.eyes_text.SetLabel(text_count(old_cursor_text)) #描写
            count = 0

        count += 1
        if count >= 30: #3秒経ったらリセット
            old_ui = ""
            old_msaa = ""
            old_cursor_text = ""
            count = 0
        sleep(0.1)

if __name__ == "__main__":
    app = wx.App(False)
    editor_frame = EditorFrame() #フレーム呼び出し
    app.SetTopWindow(editor_frame) #このフレームを最前面に
    thread1 = threading.Thread(target=get_text) #別のスレッドで実行
    thread1.setDaemon(True)
    thread1.start()
    app.MainLoop()