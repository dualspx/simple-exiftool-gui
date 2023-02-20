#IMPORTING LIBRARY
import os
from PIL import Image
from PIL.ExifTags import TAGS
from GPSPhoto import gpsphoto
import datetime
from io import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog
import olefile
import docx
##### PDF TYPE METADATA #####

# # Open the Word document
# filename = "C:/Users/User/Desktop/Degree/ITT632/ITT632-Project-CS251.doc"
# ole = olefile.OleFileIO(filename)

# # Get the metadata
# info = ole.get_metadata()

# # Print the metadata
# print(info.keys())
# # print("Author:", info.author.decode('ISO-8859-1'))
# # print("Created:", info.create_time)
# # print("Modified:", info.last_saved_time)
# # print("Subject:", info.subject.decode('utf-8'))
# # print("Title:", info.title.decode('ISO-8859-1'))

# # Close the file
# ole.close()

def extract_pdf_metadata(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)

    with open(path, 'rb') as fp:
        parser = PDFParser(fp)
        doc = PDFDocument(parser)
        return doc.info

def parse_creation_date(creation_date_str):
    # creation_date_str = creation_date_str.decode('utf-8')
    date_format = 'D:%Y%m%d%H%M%S'
    creation_date = datetime.datetime.strptime(creation_date_str[:16], date_format)
    return creation_date.strftime('%Y-%m-%d %H:%M:%S')

def parse_modification_date(mod_date_str):
    date_format = 'D:%Y%m%d%H%M%S'
    mod_date = datetime.datetime.strptime(mod_date_str[:16], date_format)
    return mod_date.strftime('%Y-%m-%d %H:%M:%S')

def pdf_metadata(file_path):
    pdf_file = file_path
    
    info = extract_pdf_metadata(pdf_file)
    creation_date_str = info[0].get('CreationDate', '')
    creation_date_str = creation_date_str.decode('utf-8')
    creation_date = parse_creation_date(creation_date_str)

    mod_date_str = info[0].get('ModDate','')
    mod_date_str = mod_date_str.decode('ISO-8859-1')
    mod_date = parse_modification_date(mod_date_str)
    
    print("THE PDF METADATA IS AT BELOW")
    for i in info:
        print('File Name\t\t :  ' + pdf_file)
        print('Author of PDF\t\t : ' + i['Author'].decode('ISO-8859-1'))
        print('Creation Date\t\t : ' + creation_date)
        if 'Creator' in i:
            print('Creator of PDF\t\t : ' + i['Creator'].decode('ISO-8859-1'))
        print('Modification Date\t : ' + mod_date)
        print('Producer\t\t : ' + i['Producer'].decode('ISO-8859-1'))


#### IMAGE TYPE METADATA #####

# path to the image or video
def image_metadata(file_path):
    imagename = file_path

    # read the image data using PIL
    image = Image.open(imagename)


    # extract other basic metadata
    info_dict = {
        "Filename": image.filename,
        "Image Size": image.size,
        "Image Height": image.height,
        "Image Width": image.width,
        "Image Format": image.format,
        "Image Mode": image.mode,
        "Image is Animated": getattr(image, "is_animated", False),
        "Frames in Image": getattr(image, "n_frames", 1)
    }

    for label,value in info_dict.items():
        print(f"{label:25}: {value}")

    # extract EXIF data
    exifdata = image.getexif()

    # iterating over all EXIF data fields
    for tag_id in exifdata:
        # get the tag name, instead of human unreadable tag id
        tag = TAGS.get(tag_id, tag_id)
        data = exifdata.get(tag_id)
        # decode bytes 
        if isinstance(data, bytes):
            data = data.decode()
        print(f"{tag:25}: {data}")

    data = gpsphoto.getGPSData(imagename)
    if(data):
        print("GPS Info\t\t :" )
        print("Latitude \t\t : ",data['Latitude'])
        print("Longitude \t\t : ", data['Longitude'])
        print("Google Map link \t : https://www.google.com/maps/search/?api=1&query="+str(data['Latitude'])+","+str(data['Longitude']))
    else:
        print('This file doesnt have GPS')
    

class App:
    def __init__(self, root):
        #setting title
        root.title("Metdata Extraction Tools")
        #setting window size
        width=637
        height=311
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_806=tk.Label(root)
        GLabel_806["cursor"] = "arrow"
        ft = tkFont.Font(family='Times',size=20)
        GLabel_806["font"] = ft
        GLabel_806["fg"] = "#333333"
        GLabel_806["justify"] = "center"
        GLabel_806["text"] = "Metadata Extractor"
        GLabel_806.place(x=0,y=20,width=631,height=46)

        enter_file_button=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        enter_file_button["font"] = ft
        enter_file_button["fg"] = "#333333"
        enter_file_button["justify"] = "center"
        enter_file_button["text"] = "Enter file"
        enter_file_button.place(x=120,y=150,width=88,height=41)

        submit_file_button=tk.Button(root)
        submit_file_button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        submit_file_button["font"] = ft
        submit_file_button["fg"] = "#000000"
        submit_file_button["justify"] = "center"
        submit_file_button["text"] = "Submit"
        submit_file_button.place(x=330,y=230,width=70,height=25)
        submit_file_button["command"] = self.submit_file_button_command

        cancel_button=tk.Button(root)
        cancel_button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        cancel_button["font"] = ft
        cancel_button["fg"] = "#000000"
        cancel_button["justify"] = "center"
        cancel_button["text"] = "Cancel"
        cancel_button.place(x=250,y=230,width=70,height=25)
        cancel_button["command"] = self.cancel_button_command

        find_file_button=tk.Button(root)
        find_file_button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        find_file_button["font"] = ft
        find_file_button["fg"] = "#000000"
        find_file_button["justify"] = "center"
        find_file_button["text"] = "Find file"
        find_file_button.place(x=470,y=150,width=62,height=38)
        find_file_button["command"] = upload_file

        file_location_placeholder=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        file_location_placeholder["font"] = ft
        file_location_placeholder["fg"] = "#333333"
        file_location_placeholder["justify"] = "center"
        self.file_location_placeholder=tk.Label(root)
        file_location_placeholder["relief"] = "sunken"
        file_location_placeholder["borderwidth"] = 2
        file_location_placeholder.place(x=200,y=150,width=265,height=36)
        # file_location_placeholder.config(text="file path")

    def submit_file_button_command(self):
        
        find_file_type()

    def cancel_button_command(self):
        file_path = " "
        print("cancel")


    def find_file_button_command():
        upload_file()
            
def show_popup_data(file_path):
    popup = tk.Toplevel()
    popup.title("Pop-up Window")
    text = tk.Text(popup, text=find_file_type())
    button = tk.Button(popup, text="OK", command=popup.destroy)
    output = image_metadata(file_path)
    text.pack(fill=tk.BOTH, expand=True)
     # Insert the output into the Text widget
    text.insert(tk.END, output)
    button.pack()

def show_popup_error():
    popup = tk.Toplevel()
    popup.title("Pop-up Window")
    label = tk.Label(popup, text="Must Insert Valid File")
    button = tk.Button(popup, text="OK", command=popup.destroy)
    label.pack(fill=tk.BOTH, expand=True)
    button.pack()

def upload_file():
    global file_path 
    file_path = filedialog.askopenfilename()

    print(file_path)
    main_window_instance.file_location_placeholder.config(text=file_path)
    # print(file_path)

def find_file_type():
    # print("File path: ",file_path)
    _, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()
    # file_path_extension = file_path.lower()

    if file_extension == '.pdf':
        # show_popup_data(file_path)
        pdf_metadata(file_path)
    elif file_extension == '.doc':
        print('doc')
    elif file_extension == '.jpg' or file_extension == '.png' or file_extension== '.jpeg' or file_extension == '.gif' or file_extension=='.jfif':
        return image_metadata(file_path)
    # add other file types as needed
    else:
        show_popup_error()


if __name__ == "__main__":
    root = tk.Tk()
    main_window_instance = App(root)
    app = App(root)
    root.mainloop()

## buat report based on the metadata
## banyakkan file type
## output ke .txt
