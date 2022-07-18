import pandas as pd
from xml.etree.ElementTree import Element, SubElement, Comment, tostring
from xml.etree import ElementTree
import time
from xml.dom import minidom
from os import listdir
from os.path import isfile, join
from zipfile import ZipFile
import xml.dom.minidom as MD

def export(exportpath,filepath,n):
    logging =""
    #Asigna etiqueta xml como prefijo.
    def prettify(elem):
        """Return a pretty-printed XML string for the Element.
        """
        rough_string = ElementTree.tostring(elem, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")

    def addValueSub(root,name,val):
        teme = SubElement(root,name)
        teme.text = val
    from datetime import datetime
    def getnamehour():
        today = datetime.now()
        st = "6_899999034_"
        return st + str(today.year) + getcomplete(str(today.month)) +getcomplete(str(today.day))+getcomplete(str(today.hour))+getcomplete(str(today.minute))+getcomplete(str(today.second))

    def getcomplete(s):
        return s if len(s)>=2 else "0"+s
    def raplaces(text):
        chars=  ["#","$","ª","º","®","º","¼","½","Á","É","Ë","Í","Ñ","Ó","×","Ú","-","'",'"','"',"™" ,"≤", "%","°",       "-","_","/","“","”","³","’"]
        carrep =["No.","","","","","","1/4","1/2","A","E","E","I","N","O","x","U","","","","","",""      , "" ," grados ","" , "", "", "", "", "", ""]
        for i in range(len(chars)):
            text = text.replace(chars[i],carrep[i])
        return text
    def escape_html(text):
        """Escape &, <, > as well as single and double quotes for HTML."""
        text = text.replace('&',"" ).replace('"',"").replace('&quot;',"").replace("&amp;","").replace("&lt;","<").replace("&gt;",">")
        return raplaces(text)
    def getinfoxml(top,rdi):
        gcl=SubElement(top, 'gcl')
        #----------------------------Deudor
        addValueSub(gcl,"foelec",str(rdi["Folio"])) 
        addValueSub(gcl,"ref",str(rdi["Referencia"])) 
        addValueSub(gcl,"canpor",'1')
    print("running...")
    logging = logging + "Running... \n"
    #Nombre del archivo a importar.
    #Si el archivo a cargar cambiar de nombre, reemplazar por el nombre actual del archivo (Ids)
    alldata = pd.read_excel("Ids.xlsx")
    alldata = alldata[pd.notna(alldata["ID"])]
    rd = alldata.to_dict(orient='records')
    for i in range(0,len(rd),n):
        rds =rd[i:i+n]
        top = Element('garantias')
        op= SubElement(top, 'op')
        addValueSub(op,"t","C") 
        addValueSub(op,"tg",str(len(rds))) 
        n=0
        for rdi in rds:
            try:
                getinfoxml(top,rdi)
                n+=1
            except Exception as e:
                print("error on " + str(rdi['ID']))
                print(str(e))
        print("File "+str(i))
        logging = logging + "File "+str(i) +"\n"

        ###################################################
        xmltext = prettify(top)
        # #Impirmir xml en consola
        #print(xmltext)
        name = exportpath+getnamehour()+".xml"
        outF = open(name, "w",encoding='utf8')
        outF.writelines(xmltext)
        print(name)
        logging = logging + name + " \n"
        print(n)
        logging = logging + str(n) + " \n"
        outF.close()
        time.sleep(11)      
        #####################################################
    return logging
import wx

class OtherFrame(wx.Frame):
    """
    Class used for creating frames other than the main one
    """
    def __init__(self, title, parent=None):
        wx.Frame.__init__(self, parent=parent, title=title,pos=(60,60))
        self.panel = wx.Panel(self)
        self.my_sizer = wx.BoxSizer(wx.VERTICAL)
    def print_on_frame(self,text):
        textelement = wx.StaticText(self.panel)
        textelement.SetLabel(text)
        self.my_sizer.Add(textelement, 0, wx.ALL | wx.EXPAND, 5)
        self.panel.SetSizer(self.my_sizer)
        self.Show()

class MyFrame(wx.Frame):    
    def __init__(self):
        super().__init__(parent=None, title='Generador de XMl')
        panel = wx.Panel(self)        
        my_sizer = wx.BoxSizer(wx.VERTICAL)        
        self.text_ctrl = wx.TextCtrl(panel)
        my_sizer.Add(self.text_ctrl, 0, wx.ALL | wx.EXPAND, 5)        
        my_btn = wx.Button(panel, label='Selecionar Archivo de IDs')
        my_btn.Bind(wx.EVT_BUTTON, self.on_press)
        my_sizer.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)
        self.name_ctrl = wx.TextCtrl(panel)
        self.name_ctrl.SetValue("Nombre archivo exportar")
        my_sizer.Add(self.name_ctrl, 0, wx.ALL | wx.EXPAND, 5)
        self.n_ctrl = wx.TextCtrl(panel)
        self.n_ctrl.SetValue("Numero de planes por XML")
        my_sizer.Add(self.n_ctrl, 0, wx.ALL | wx.EXPAND, 5)
        btn_process = wx.Button(panel, label='Process')
        btn_process.Bind(wx.EVT_BUTTON, self.process)
        my_sizer.Add(btn_process, 0, wx.ALL | wx.CENTER, 5)
        # set sizeer
        panel.SetSizer(my_sizer)
        
        # show interface
        self.Show()
    def on_press(self,event):
        # Create open file dialog
        openFileDialog = wx.FileDialog(frame, "Open", "", "", 
            "*", 
            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        self.text_ctrl.SetValue (openFileDialog.GetPath())
    def process(self,event):
        #namefile : Nombre del archivo a exportar
        path = self.text_ctrl.GetValue()
        namefile=self.name_ctrl.GetValue()
        exportpath = ".\export\ "
        #folder = .\Data-plan\
        filepath = ".\Data-plan\ "+namefile+".xlsx"
        n = int(self.n_ctrl.GetValue())
        logs=""
        try:
            #logs += cross(path,namefile)
            logs+= export(exportpath,filepath,n)
        except Exception as e:
            log = str(e)
            print(e)
        print("LOG",logs)
        self.frame = OtherFrame(title="logging")
        self.frame.print_on_frame(logs)
        
if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()