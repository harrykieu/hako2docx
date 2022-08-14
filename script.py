import os
import re
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.shared import RGBColor
import PIL
from PIL import Image
import validators


def strip_url(URL):
    global lnname, mainsoup
    # Get main URL and main ln name
    mainURL = re.sub('\/c.*$','',URL) #get the main url of ln
    mainpage = requests.get(mainURL)
    mainsoup = BeautifulSoup(mainpage.content,'html.parser')
    lnname = mainsoup.find("span",attrs='series-name').text.strip()

def find_volume():
    global volURL, baseURL
    
    baseURL = 'https://ln.hako.vn'
    volURL_wobase = mainsoup.find_all("a",href=True)
    volnum = 1
    volURLdict = {}
    volnamedict = {}
    print("Get the volume:")
    for x in volURL_wobase:
        if (re.search('/t[0-9]',x['href']))!= None:
            volURL = baseURL + x['href']
            volweb = requests.get(volURL)
            volsoup = BeautifulSoup(volweb.content,'html.parser')
            volname = volsoup.find("span",attrs='volume-name').text.strip()
            print(f'{volnum}. {volname}')
            volURLdict[volnum] = volURL
            volnamedict[volnum] = volname
            volnum +=1
    volinp = ""
    while volinp not in range(volnum):
        try:
            print("Choose volume:\n>", end=" ")
            volinp = int(input())
        except ValueError:
            print("Invalid volume. Please enter again:")
            pass

    print(f"Chosen the volume: {volnamedict[volinp]}")
    volURL = volURLdict[volinp]
    print(f"URL: {volURL}")

#----------------------------------------------------------

def get_URL_chap():
    global chapURL
    chapweb = requests.get(volURL)
    chapsoup = BeautifulSoup(chapweb.content,'html.parser')
    chapURL_raw = chapsoup.find_all("a",href=True)
    chapURL = []
    for x in chapURL_raw:
        if (re.search('/c[0-9]',x['href']))!= None:
            chapURL.append(baseURL+x['href'])
    
#----------------------------------------------------------

def find_contents(URL):
    
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser")
    
    # Pull the content from web
    lncontents = soup.find_all("p",id=True)
    volname = soup.find("h2",attrs="title-item").contents[0].text
    title = soup.find("h4",attrs="title-item").contents[0].text
    notes = soup.find_all(attrs="note-content long-text")
    

    # Create an instance of Document
    document = Document()

    # Edit styles
    styles = document.styles
    styles['Heading 1'].font.color.rgb = RGBColor(0, 0, 0)
    styles['Heading 1'].font.size = Pt(14)
    styles['Normal'].font.name = "Cambria"
    styles['Normal'].font.color.rgb = RGBColor(0,0,0)
    styles['Normal'].font.size = Pt(12)

    # Add heading
    head = document.add_heading(title,1)
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Default value 
    img_num = 1  
    note_num = 1
    
    # Make the default folder
    d = os.path.dirname(__file__) # directory of script
    if "..." in title:
        pathname = f'{d}/{lnname}/{volname}/{title.replace("...","")}'   
    else:
        pathname = f'{d}/{lnname}/{volname}/{title}'
    imgfolder = f'{pathname}/image'
    os.makedirs(os.path.dirname(imgfolder),exist_ok=True)

    # Add contents
    for cnt in lncontents:
        for cont in cnt.contents:
            # Make the image file
            imgname = f'{imgfolder}/img{img_num}.jpg'
            os.makedirs(os.path.dirname(imgname), exist_ok=True)
            
            # Find the image
            if cont.name == 'img':
                img_url = cont['src']
                img_data = requests.get(img_url).content 
                
                # Get the image
                try:
                    with open(imgname,'wb+') as img:
                        img.write(img_data)
                    img.close()
                    imge = PIL.Image.open(imgname)
                    imge.save(imgname)
                
                # Delete the image file if can't get the file
                except PIL.UnidentifiedImageError:
                    document.add_paragraph(f"img{img_num}.jpg (get the image by hand)")
                    os.remove(imgname)

                # Save the image
                else:
                    document.add_paragraph(f"img{img_num}.jpg (on the folder)")
                    img_num += 1

        # Add paragraph for text + some alignment
        parrg = document.add_paragraph(style='Normal')
        parrg.paragraph_format.line_spacing = 1.0
        parrg.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY        
        
        # Add text in paragraph word by word
        for word in cnt:
            run = parrg.add_run() # Allow making word Bold, Italic, Strikethrough
            strip_text=word.text.replace(u'\u200e','')
            
            # Search for [note...] and replace
            test = re.search('\[note[0-9][0-9][0-9][0-9][0-9]\]',word.text) # Regex for [note...]
            if test!= None:
                k = word.text.replace(test.group(0),f" (Note {note_num})")
                note_num+=1
                run.add_text(k)
            else:
                run.add_text(strip_text)
            
            # Making word Bold, Italic, Strikethrough
            if word.name == 's':
                run.font.strike = True
                continue
            elif word.name == 'strong':
                run.bold = True
                continue
            elif word.name == 'em':
                run.italic = True
                continue
    
    # Add notes at the end of document
    if notes:
        document.add_paragraph("\nNote:")
        stt=1
        for note in notes:
            nte = note.text.strip()
            document.add_paragraph(f"{stt}. {nte}",style='Normal')
            stt+=1

    # Remove the 'image' folder if there is no image            
    if not os.listdir(imgfolder):
        os.rmdir(imgfolder)
    
    # Save the document        
    document.save(f"{pathname}/{title}.docx")

#---------------------------------------------------------------------------------------
# User interface
print("+------------------------------------------------------+")
print("|                     HAKOTODOCX                       |")
print("+------------------------------------------------------+")
print("Please enter your choice:")
inp = ""
while (inp not in ["1","2","3"]):
    print("1: Get a whole volume\n2: Get a chapter\n3: Exit\n>",end=" ")
    inp = input()
    if (inp not in ["1","2","3"]):
        print("Invalid. Please enter your choice: ")
        continue
    elif inp=="1":
        print("Please enter your URL:")
        print(">",end=" ")
        URL = input()
        while not validators.url(URL):
            print("Invalid URL. Please enter URL:")
            URL = input()
        else:
            strip_url(URL)
            find_volume()
            get_URL_chap()
            for x in chapURL:
                find_contents(x)
            print("Done! Please enter your choice:")
            inp = ""
    elif inp=="2":
        print("Please enter your URL:")
        print(">",end=" ")
        URL = input()
        while not validators.url(URL):
            print("Invalid URL. Please enter URL:")
            URL = input()
        else:
            strip_url(URL)
            find_contents(URL)
            print("Done! Please enter your choice:")
            inp = ""
    elif inp=="3":
        break
