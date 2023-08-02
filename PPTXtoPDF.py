#https://gist.github.com/aspose-com-gists/cbc95264bf51d082800eca055e38cc5c

import shutil
import os
import win32com.client
import time



def move_converted_slides(src_folder,dst_folder, file):
    try:
        if not os.path.exists(dst_folder):
            os.makedirs(dst_folder)
            
        if os.path.exists(dst_folder + file):
            os.remove(src_folder)
        elif not os.path.exists(dst_folder + file):
            shutil.move(src_folder, dst_folder)
    except:
        return
    



def pptx_to_pdf(inputFileName, outputFileName, formatType = 32):
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName, WithWindow=False)
    deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()


def search_for_pptx_files(folder_path):
    try:
        for root, dirs, files in os.walk(folder_path):
            if "SlidesBackup" in root:
                continue # Skip this folder
            for file in files:
                if file.endswith(".pptx"):
                    if (file[:2] == "~$"):
                        pptx_to_pdf(root+"\\"+file[2:], root+"\\"+file[2:-5])
                        move_converted_slides(root+"\\"+file[2:], root+"\\SlidesBackup\\", file[2:])
                    else:
                        pptx_to_pdf(root+"\\"+file, root+"\\"+file[:-5])
                        move_converted_slides(root+"\\"+file, root+"\\SlidesBackup\\", file)
    except Exception as ex:
        print(ex)


while True:
    search_for_pptx_files("C:\\Users\\jeand\\My Drive\\University")
    time.sleep(5)
    # no need for 1 second, changed it to 5


