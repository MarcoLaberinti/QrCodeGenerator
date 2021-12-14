
import qrcode
from tkinter import *
from tkinter import filedialog
from PIL import Image, ImageFont, ImageDraw, ImageTk

root = Tk()
root.geometry("680x430")
root.resizable(False, False)
root.title("QR Code Generator")
root.iconbitmap('icona_1.ico')

C = Canvas(root, bg="blue", height=250, width=300)
filename = PhotoImage(file = "background.png")
background_label = Label(root, image=filename)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
#C.pack()



##--- qr code setting ---##
qr=qrcode.QRCode(
    version=1,
    box_size=8,
    border=5
    )
##--- function ---##
folder_selected=""
flag = StringVar()

def genera():
    
    name=name_field.get()
    allied=allied_field.get()
    tectubi=tectubi_field.get()
    riga12=riga12_field.get()
    flag1=(flag.get())
    
    if flag1!="OFF":

        val_iniziale=str(val_iniziale_field.get())
        val_finale=str(val_finale_field.get())
        
        val_iniziale_int=int(val_iniziale)
        val_finale_int=int(val_finale)
        
        i=val_iniziale_int
        j=1
        
        num=val_finale_int-val_iniziale_int
        while i<=val_finale_int:
            
            qr=qrcode.QRCode(
                version=1,
                box_size=10,
                border=7
                )
              
            qr.add_data(tectubi + riga12 + allied+str(i))
            qr.make(fit=True)
            
            img=qr.make_image(fill="black", back_color="white")
            
            RImg=img.resize((302, 302))
            width, height = RImg.size
            draw = ImageDraw.Draw(RImg)
            
            font = ImageFont.truetype("arialbd.ttf", size=20)
            draw.text((45, height - 25), allied+str(i), font=font)
            draw.text((45, height - 48), tectubi, font=font)
            draw.text((195,height - 48), riga12, font=font)

            RImg.save(folder_selected + "/" + name + str(i)+".png", "png")

            success = Label(root, text="Codice generato! "+str(j-1)+"/"+str(num)).grid(row=12, column=3)
            i=int(i)+1
            j=int(j)+1
    else:
        qr=qrcode.QRCode(
            version=1,
            box_size=10,
            border=7
            )
          
        qr.add_data(tectubi + riga12 + allied)
        qr.make(fit=True)
        
        img=qr.make_image(fill="black", back_color="white")
        
        RImg=img.resize((302, 302))
        width, height = RImg.size
        draw = ImageDraw.Draw(RImg)
        
        font = ImageFont.truetype("arialbd.ttf", size=20)
        draw.text((45, height - 25), allied, font=font)
        draw.text((45, height - 48), tectubi, font=font)
        draw.text((195,height - 48), riga12, font=font)

        RImg.save(folder_selected + "/" + name +".png", "png")

        success = Label(root, text="  Codice generato! 1/1  ").grid(row=12, column=3)
            
        

    name_field.delete(0, END)
    allied_field.delete(0,END)
    tectubi_field.delete(0, END)
    riga12_field.delete(0, END)
    val_iniziale_field.delete(0, END)
    val_finale_field.delete(0, END)
    

def directory():
    global folder_selected
    folder_selected = filedialog.askdirectory()
    return folder_selected



##--- checkbox ---##

enable_series = Checkbutton(root, text="", variable=flag, onvalue="ON", offvalue="OFF").grid(row=10, column=2)

##--- text field ---##
name_field = Entry(root, width=60, textvariable="nome")

tectubi_field = Entry(root,width=25)

riga12_field = Entry(root, width=25)


allied_field = Entry(root,width=60)

val_iniziale_field = Entry(root, width=15)
val_finale_field = Entry(root, width=15)




##--- tkinter label ---##


name_file = Label(root, text="       Inserisci il nome con cui vuoi chiamare il file: ")

tectubi_text = Label(root, text="Prima riga: ")

allied_text = Label(root, text="Seconda riga: ")

intervallo_text = Label(root, text="Intervallo: ")

    
save_folder = Label(root, text="Scegli la cartella del salvataggio")

space11 = Label(root, text="\n\n  ").grid(row=1, column=1)
space31 = Label(root, text="  ").grid(row=3, column=1)
space51 = Label(root, text="  ").grid(row=5, column=1)
space71 = Label(root, text="  ").grid(row=7, column=1)
space94 = Label(root, text="  ").grid(row=9, column=4)
space111 = Label(root, text="  ").grid(row=11, column=1)
space134 = Label(root, text="  ").grid(row=13, column=4)


##--- button ---##

save_button = Button(root, text="Genera",command=genera, fg="#536D3B", padx=20, pady=10,)# state=DISABLE, padx=50, pady=50, fg="#536D3B", bg="#FF5733")

browse_button = Button(root, text="Browse", command=directory, fg="#536D3B")
    

## --- grid ---##


name_file.grid(row=2, column=2, sticky="E")
name_field.grid(row=2, column=3)

save_folder.grid(row=4, column=2, sticky="E")

browse_button.grid(row=4, column=3, sticky="W")

tectubi_text.grid(row=6, column=2, sticky="E")
tectubi_field.grid(row=6, column=3, sticky="W")
riga12_field.grid(row=6, column=3, sticky="E")

allied_text.grid(row=8, column=2, sticky="E")
allied_field.grid(row=8, column=3)

intervallo_text.grid(row=10, column=2,sticky="E")
val_iniziale_field.grid(row=10, column=3, sticky="W")
val_finale_field.grid(row=10, column=3)



save_button.grid(row=12, column=3, sticky="E")

load= Image.open("foreground1.png")
render = ImageTk.PhotoImage(load)
img = Label(root, image=render)
img.place(x=13, y=30)


root.mainloop()

    
