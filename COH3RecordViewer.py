import time
from datetime import datetime
import shutil
import os
import sys
import winreg
from tkinter import *
import threading
from tkscrolledframe import ScrolledFrame
import ctypes.wintypes
from PIL import Image, ImageTk
import pyperclip
import webbrowser
import psutil
import subprocess

##Permet de récupérer l'adresse du dossier "mes documents"
CSIDL_PERSONAL = 5       # My Documents
SHGFP_TYPE_CURRENT = 0   # Get current, not default value

buf= ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
##Permet de récupérer l'adresse du dossier "mes documents"


class ToolTip(object): #https://stackoverflow.com/questions/20399243/display-message-when-hovering-over-something-with-mouse-cursor-in-python

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 30
        y = y + cy + self.widget.winfo_rooty() +25
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

def returnplayer(line): #Renvoie le nom du joueur et sa nation
    words = line.split()
    player = [' '.join(words[8:-3]),words[-1]]
    return player

def returnmap(line):
    words = line.split('\\')
    return words[-1]
    
def Rename_file(event):
    if entry.get() == "":
        warning = Label(win_rename, text="Le nom du fichier ne peut être vide",fg='red', font=("Calibri", 11))
        warning.grid(row = 3, padx=10, sticky = "W")
        win_rename.update()
        win_rename.geometry(str(win_rename.grid_bbox()[2])+"x"+str(win_rename.grid_bbox()[3]))
    else :
        new_file_name = entry.get()
        old_path_rec = playbak_path+"\\"+old_file_name_too
        old_path_dat = playbak_path+"\\"+old_file_name_too[:-4]+".txt"
        new_path_rec = playbak_path+"\\"+new_file_name+".rec"
        new_path_dat = playbak_path+"\\"+new_file_name+".txt"
        os.rename(old_path_rec, new_path_rec)
        os.rename(old_path_dat, new_path_dat)
        win_rename.destroy()
        update_window()

def display_text_box(old_file_name):
    global win_rename
    global entry
    global old_file_name_too 
    main_window_coordinates_x = window.geometry().split('+')[1]
    main_window_coordinates_y = window.geometry().split('+')[2]
    
    old_file_name_too = old_file_name
    new_file_name = ""
    win_rename = Tk()
    win_rename.title('Renommer le replay')
    entry = Entry(win_rename, width = 50)
    entry.insert(0, old_file_name[:-4])
    # Association de l'évènement actionEvent au champ de saisie
    info = Label(win_rename, text="Nom du nouveau fichier :", font=("Calibri", 11))
    info.grid(row = 1, sticky = "W", padx=10)
    entry.bind("<Return>", Rename_file) 
    entry.grid(row=2, padx=10, pady = 5) 
    win_rename.update()
    text_box_width = win_rename.grid_bbox()[2]
    text_box_height = win_rename.grid_bbox()[3]
    text_box_pos_x = int(int(main_window_coordinates_x)+inner_frame.grid_bbox()[2]/2-text_box_width/2)
    text_box_pos_y = int(int(main_window_coordinates_y)+inner_frame.grid_bbox()[3]/2-text_box_height/2)
    win_rename.geometry(str(text_box_width)+"x"+str(text_box_height)+'+'+str(text_box_pos_x)+'+'+str(text_box_pos_y))
    # win_rename.geometry("300x300+1000+0")
    win_rename.mainloop()
    

def analyse(warnings_path, playbak_path):
    global COHrunning
    global thread_on
    global acquisition_on
    isrecord = 0
    maps = ''
    waitforcoh()
    if COHrunning == 1:
        with open(warnings_path,'r', encoding = "UTF-8", errors='replace') as f:
            while isrecord == 0 and mainwindow_open == True and var.get() == 1:
                for line in f:
                    if 'REC --' in line:
                        isrecord = 1
                        print("Replay en cours, acquisition annulée")
                        btn1.deselect()
                        COHrunning = 0
                        break
                    if 'GAME -- Scenario' in line:
                        joueurs = []
                        maps = returnmap(line)
                        print("Partie détectée avec ces paramètres :")
                        print("   "+line.replace("\n",""))
                        start_time = time.time()
                    if 'GAME -- Human Player' in line or 'GAME -- AI Player' in line:
                        joueurs.append(returnplayer(line))
                        print("   "+line.replace("\n",""))
                    if 'GameObj::ShutdownGameObj' in line:
                        end_time = time.time()
                        if end_time-start_time > 10:
                            print("Partie terminée")
                            shutil.copyfile(playbak_path+'\\temp.rec',playbak_path+'\\'+datetime.now().strftime("%d-%m-%Y_%Hh%Mm")+'.rec')
                            o = open(playbak_path+'\\'+datetime.now().strftime("%d-%m-%Y_%Hh%Mm")+'.txt', 'w', encoding = "UTF-8", errors='replace')
                            o.write(maps)
                            o.write(time.strftime("%H:%M:%S", time.gmtime(end_time-start_time))+"\n")
                            for i in joueurs:
                                o.write(i[0]+' '+i[1]+'\n')
                            o.close()
                            addbuttons()   
                        else:
                            print("Il s'agit certainement d'une partie précédente")
                            isrecord = 0
                            joueurs = []
                    if 'Application closed without errors' in line:
                        print("Jeu fermé")
                        btn1.deselect()
                        COHrunning = 0
                time.sleep(1)
            f.close()
    thread_on = 0
    acquisition_on = 0
    print("Fin de l'acquisition")


def Thread_analyse(warnings_path, playbak_path):
    global thread_on
    global acquisition_on
    if var.get() == 1 and thread_on == 0:
        print("Acquisition lancée")
        acquisition_on = 1
        thread_on = 1
        thread = threading.Thread(target=analyse, args=[warnings_path, playbak_path])
        thread.start()

def listerecords(playbak_path):
    list = []
    for path in os.listdir(playbak_path):
        # check if current path is a file
        if os.path.isfile(os.path.join(playbak_path, path)) and path!='temp.rec' and path!='temp_campaign.rec' and path[-4:]=='.rec':
            list.append(path)
    return list

def launchrecord(filename):
    if var2.get() == 1:
        pyperclip.copy("dofile('replay-enhancements/init.scar')")
    home_path = os.getcwd()
    hkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\WOW6432Node\Valve\Steam")
    steam_path = winreg.QueryValueEx(hkey, "InstallPath")[0]
    winreg.CloseKey(hkey)
    os.chdir(steam_path)
    os.system(".\steam.exe -applaunch 1677280 -dev -replay playback:"+filename)
    os.chdir(home_path)
    
def removerecord(i, playbak_path):
    os.remove(playbak_path+"\\"+i)
    try:
        os.remove(playbak_path+"\\"+i[:-4]+".txt")
    except:
        print("Le fichier data n'existe pas")
    update_window()

def update_window():
    global iterate
    global game_number
    global mapspiclist
    global var2
    global replay_mod_on
    global main_window_coordinates_x
    global main_window_coordinates_y    
    main_window_coordinates_x = window.geometry().split('+')[1]
    main_window_coordinates_y = window.geometry().split('+')[2]  
    iterate=0
    game_number = 0
    mapspiclist = []
    replay_mod_on = var2.get()
    window.destroy()
    generatewindow()
    

def addbuttons():
    global iterate
    global game_number
    global maps_icon
    global main_window_coordinates_x
    global main_window_coordinates_y    
    linepositions= []
    for i in listerecords(playbak_path)[game_number:]:
        try:
            info = open(playbak_path+"\\"+i[:-4]+".txt", "r", encoding = "UTF-8", errors='replace')
        except:
            linenumberingrid = 1
            iterate+=1
            mapspiclist.append([])
            for j in range(2,9):
                if j != 5 and j != 6:
                    label=Label(inner_frame, text='N.A.', font='calibri 9')
                    label.grid(row=5+iterate-linenumberingrid,column=j, rowspan=linenumberingrid, padx=20,pady=15)
        else:
            number_allies = 0
            number_axes = 0
            linenumberingrid = 0
            lineinfile = 1
            mapname=''
            gamelength = ''
            for line in info:
                if lineinfile == 1:
                    mapname = line.replace("\n","")
                elif lineinfile == 2:
                    gamelength = line.replace("\n","")
                elif lineinfile > 2:
                    label=Label(inner_frame, text=' '.join(line.split()[:-1]), font='calibri 9')
                    if line.split()[-1] == "british_africa" or line.split()[-1] == "americans":
                        label.grid(row=5+number_allies+iterate,column=4, sticky = E)
                        if line.split()[-1] == "british_africa":
                            pic = Label(inner_frame, image = UKF_icon)
                            pic.grid(row=5+number_allies+iterate,column=5, sticky = W)
                        if line.split()[-1] == "americans":
                            pic = Label(inner_frame, image = USF_icon)
                            pic.grid(row=5+number_allies+iterate,column=5, sticky = W)
                        number_allies += 1
                        if number_allies > number_axes:
                            linenumberingrid += 1
                    else:
                        label.grid(row=5+number_axes+iterate,column=7, sticky = W)
                        if line.split()[-1] == "germans":
                            pic = Label(inner_frame, image = WM_icon)
                            pic.grid(row=5+number_axes+iterate,column=6, sticky = E)
                        if line.split()[-1] == "afrika_korps":
                            pic = Label(inner_frame, image = DAK_icon)
                            pic.grid(row=5+number_axes+iterate,column=6, sticky = E)
                        number_axes += 1
                        if number_axes > number_allies:
                            linenumberingrid += 1
                lineinfile += 1
            iterate += max(number_allies, number_axes)
            
            mapspic = Image.open(".\\Assets\\Maps\\"+mapname+".png")
            resized_maps = mapspic.resize((150,150), Image.LANCZOS)
            mapspic.close()
            mapspiclist.append(ImageTk.PhotoImage(resized_maps))

            labellenght = Label(inner_frame, text=gamelength, font='calibri 10')
            labellenght.grid(row=5+iterate-linenumberingrid, column=3, rowspan=linenumberingrid)

            labelmap = Label(inner_frame, image = mapspiclist[game_number])
            labelmap.grid(row=5+iterate-linenumberingrid,column=8, rowspan=linenumberingrid, padx=15)
            
            info.close()

        btnsuppr= Button(inner_frame, image = Delete_icon, text="X", font=("Calibri", 20, "bold"),fg='red', command= lambda i=i: removerecord(i, playbak_path), anchor="center", relief='flat')
        btnsuppr.grid(row=5+iterate-linenumberingrid,column=9, rowspan=linenumberingrid, padx=15,pady=15)
        
        btnfile= Button(inner_frame, text=str(i), font=("Calibri", 13),command= lambda i=i: launchrecord(i), anchor="center")
        btnfile.grid(row=5+iterate-linenumberingrid,column=2, rowspan=linenumberingrid, padx=20,pady=15, sticky = "W")
        
        renamefile= Button(inner_frame, image = Pen_icon, font=("Calibri", 13),command= lambda filename=i: display_text_box(filename), anchor="center", relief='flat')
        renamefile.grid(row=5+iterate-linenumberingrid,column=1, rowspan=linenumberingrid, sticky = "E")
        CreateToolTip(renamefile, text = 'Renommer')
        
        window.update()

        linepositions += [5+iterate]  
        for i in linepositions:
            w = Canvas(inner_frame, width = inner_frame.grid_bbox()[2]-20, height = 1, bg='black')
            w.grid(row=i,column=1, columnspan=9)   
        iterate += 1
        game_number += 1

    if listerecords(playbak_path)[game_number:] == []:
        window.update()
    try:
        window.geometry(str(inner_frame.grid_bbox()[2]+20)+"x600+"+main_window_coordinates_x+'+'+main_window_coordinates_y) #Définit la taille de la fenêtre
    except:
        window.geometry(str(inner_frame.grid_bbox()[2]+20)+"x600") #Définit la taille de la fenêtre

def callback(url):
   webbrowser.open_new_tab(url)

def aboutwindow(): 
    main_window_coordinates_x = window.geometry().split('+')[1]
    main_window_coordinates_y = window.geometry().split('+')[2]
    
    win_about = Tk()
    win_about.title('A propos')
    
    name_label=Label(win_about, text="COH 3 Record Viewer", font=("Calibri", 20))
    name_label.grid(row=1,column=1, padx=30)
    author_label=Label(win_about, text="Auteur : David Germain", font=("Calibri", 11))
    author_label.grid(row=2,column=1, padx=30)
    contribute_label=Label(win_about, text="\nContributions :", font=("Calibri", 16))
    contribute_label.grid(row=3,column=1, padx=30)
    contribute_name1_label=Label(win_about, text="        • squareRoot17", font=("Calibri", 11))
    contribute_name1_label.grid(row=4,column=1, padx=30, sticky = W)
    contribute_name2_label=Label(win_about, text="        • ewerybody", font=("Calibri", 11))
    contribute_name2_label.grid(row=5,column=1, padx=30, sticky = W)
    
    other_software_label=Label(win_about, text="Autres logiciels :", font=("Calibri", 16))
    other_software_label.grid(row=6,column=1, padx=30)    
    
    COH_replay_enhancer_label=Label(win_about, text="coh3-replay-enhancements :", font=("Calibri", 13))
    COH_replay_enhancer_label.grid(row=7,column=1, padx=30, sticky = W)
    COH_replay_enhancer_author_label=Label(win_about, text="        • auteur : Janne252", font=("Calibri", 11))
    COH_replay_enhancer_author_label.grid(row=8,column=1, padx=30, sticky = W)
    COH_replay_enhancer_link_label=Label(win_about, text="        • https://github.com/Janne252/coh3-replay-enhancements",font=('Calibri', 11), fg="blue", cursor="hand2")
    COH_replay_enhancer_link_label.grid(row=9,column=1, padx=30, sticky = W)
    COH_replay_enhancer_link_label.bind("<Button-1>", lambda e: callback("https://github.com/Janne252/coh3-replay-enhancements"))
    win_about.update()
    
    about_box_width = win_about.grid_bbox()[2]
    about_box_height = win_about.grid_bbox()[3]
    about_box_pos_x = int(int(main_window_coordinates_x)+inner_frame.grid_bbox()[2]/2-about_box_width/2)
    about_box_pos_y = int(int(main_window_coordinates_y)+inner_frame.grid_bbox()[3]/2-about_box_height/2)
    
    win_about.geometry(str(about_box_width)+"x"+str(about_box_height)+'+'+str(about_box_pos_x)+'+'+str(about_box_pos_y))
    win_about.mainloop()

def generatewindow():
    global acquisition_on
    global replay_mod_on
    global inner_frame
    global window
    global var
    global var2
    global sf
    global btn1
    global btn2
    global UKF_icon
    global USF_icon
    global WM_icon
    global DAK_icon
    global Delete_icon
    global Pen_icon
    window=Tk()
    window.iconbitmap("./Assets/Misc/icon.ico")
    window.title('COH3 replay viewer UI')
    
    UKF = Image.open(".\\Assets\\Factions\\UKF.png")
    resized_UKF = UKF.resize((20,20), Image.LANCZOS)
    UKF_icon = ImageTk.PhotoImage(resized_UKF)

    USF = Image.open(".\\Assets\Factions\\USF.png")
    resized_USF = USF.resize((20,20), Image.LANCZOS)
    USF_icon = ImageTk.PhotoImage(resized_USF)
    
    WM = Image.open(".\\Assets\Factions\\WM.png")
    resized_WM = WM.resize((20,20), Image.LANCZOS)
    WM_icon = ImageTk.PhotoImage(resized_WM)
    
    DAK = Image.open(".\\Assets\Factions\\DAK.png")
    resized_DAK = DAK.resize((20,20), Image.LANCZOS)
    DAK_icon = ImageTk.PhotoImage(resized_DAK)
    
    Refresh = Image.open(".\\Assets\\Misc\\Refresh.png")
    resized_Refresh = Refresh.resize((30,30), Image.LANCZOS)
    Refresh_icon = ImageTk.PhotoImage(resized_Refresh)
    
    Delete = Image.open(".\\Assets\\Misc\\delete.png")
    resized_Delete = Delete.resize((40,40), Image.LANCZOS)
    Delete_icon = ImageTk.PhotoImage(resized_Delete)
    
    Folder = Image.open(".\\Assets\\Misc\\Folder.png")
    resized_Folder = Folder.resize((30,30), Image.LANCZOS)
    Folder_icon = ImageTk.PhotoImage(resized_Folder)
    
    Pen = Image.open(".\\Assets\\Misc\\Pen.png")
    resized_Pen = Pen.resize((30,30), Image.LANCZOS)
    Pen_icon = ImageTk.PhotoImage(resized_Pen)
    
    question_mark = Image.open(".\\Assets\\Misc\\Question_mark.png")
    resized_question_mark = question_mark.resize((30,30), Image.LANCZOS)
    question_mark_icon = ImageTk.PhotoImage(resized_question_mark)
    
    var = IntVar(value = acquisition_on) #enregistre la valeur du bouton d'acquisition
    var2 = IntVar(value = replay_mod_on) #enregistre la valeur du bouton de copie
    
    sf = ScrolledFrame(window)
    sf.pack(side="top", expand=1, fill="both")
    sf.bind_arrow_keys(window)
    sf.bind_scroll_wheel(window)
    inner_frame = sf.display_widget(Frame)

    upd = Button(inner_frame, image = Refresh_icon, text="update", font=("Calibri", 10, "bold"), command= lambda : update_window(), relief='flat')
    upd.grid(row=1,column=1, sticky = W)
    CreateToolTip(upd, text = 'Mettre à jour la liste des fichiers')
    
    open_folder = Button(inner_frame, image = Folder_icon, text="Ouvrir dossier des playbacks", font=("Calibri", 10), command= lambda : subprocess.Popen('explorer '+playbak_path), relief='flat')
    open_folder.grid(row=1,column=2, sticky = W)
    CreateToolTip(open_folder, text = 'Ouvrir le dossier des playbacks')

    about_btn = Button(inner_frame, image = question_mark_icon, text="A propos", font=("Calibri", 10), command= aboutwindow, relief='flat')
    about_btn.grid(row=1,column=9, sticky = E)
    CreateToolTip(about_btn, text = 'A propos')

    lbl=Label(inner_frame, text="Company of Heroes 3 Replay Viewer", fg='red', font=("Calibri", 16))
    lbl.grid(row=1,column=1, columnspan=9)

    btn1=Checkbutton(inner_frame, text="Lancer l'acquisition", font=("Calibri", 13), variable=var, command= lambda : Thread_analyse(warnings_path, playbak_path), anchor="center")
    btn1.grid(row=2,column=1, columnspan=4)

    btn2=Checkbutton(inner_frame, text="""Mode 'COH3-Replay-Enhancement'""", font=("Calibri", 13), variable=var2, anchor="center")
    btn2.grid(row=2,column=6, columnspan=4)
    CreateToolTip(btn2, text = """Copie la commande 'dofile('replay-enhancements/init.scar')' dans le presse papier lors du lancement d'un replay.\nUtile si combiné à l'utilisation du mod 'coh3-replay-enhancements' de Janne252""")


    Col1=Label(inner_frame, text="Fichier", fg='red', font=("Calibri", 14), anchor="center")
    Col1.grid(row=3,column=1, columnspan = 2)
    
    Col2=Label(inner_frame, text="Durée", fg='red', font=("Calibri", 14), anchor="center")
    Col2.grid(row=3,column=3)
    
    Col3=Label(inner_frame, text="Alliés", fg='red', font=("Calibri", 14), anchor="center")
    Col3.grid(row=3,column=4, columnspan=2, padx=15)
    
    Col4=Label(inner_frame, text="Axe", fg='red', font=("Calibri", 14), anchor="center")
    Col4.grid(row=3,column=6, columnspan=2, padx=15)
    
    Col5=Label(inner_frame, text="Map", fg='red', font=("Calibri", 14), anchor="center")
    Col5.grid(row=3,column=8, padx=15)
    
    Col6=Label(inner_frame, text="Supprimer", fg='red', font=("Calibri", 14), anchor="center")
    Col6.grid(row=3,column=9, padx=15)
    
    addbuttons()
    
    w = Canvas(inner_frame, width = inner_frame.grid_bbox()[2]-20, height = 1, bg='black')
    
    w.grid(row=4,column=1, columnspan=9)
    
    if acquisition_on == 1 and thread_on == 0:
        Thread_analyse(warnings_path, playbak_path)
    
    # window.update()
    # window.geometry(str(inner_frame.grid_bbox()[2]+20)+"x600") #Définit la taille de la fenêtre

    window.resizable(False, False)
    window.protocol("WM_DELETE_WINDOW", close_window)
    window.mainloop()

def process_exists(process_name): #https://stackoverflow.com/questions/7787120/check-if-a-process-is-running-or-not-on-windows
    return process_name in (p.name() for p in psutil.process_iter())

def waitforcoh():
    global COHrunning
    if process_exists('RelicCoH3.exe') == False:
        COHrunning = 0
    if COHrunning == 0 :
        while process_exists('RelicCoH3.exe') == False and mainwindow_open == True and var.get() == 1:
            time.sleep(1)
        if var.get() == 1 and mainwindow_open == True:
            COHrunning = 1
            print("Jeu détecté")
            time.sleep(5)

def close_window():
    global mainwindow_open
    mainwindow_open = False  # turn off while loop
    print("Logiciel fermé")
    window.destroy()
    exit()
   

playbak_path = buf.value + "\\My Games\\Company of Heroes 3\\playback" #Contient l'emplacement du dossier qui contient les playbacks
warnings_path = buf.value + "\\My Games\\Company of Heroes 3\\warnings.log" #Contient l'emplacement du dossier qui contient les playbacks
iterate = 0
game_number=0
COHrunning = 0
mapspiclist = []
thread_on = 0
acquisition_on = 1
replay_mod_on = 1
mainwindow_open = True

generatewindow()