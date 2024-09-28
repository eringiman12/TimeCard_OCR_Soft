import glob
import os
import shutil

File_path = r"処理後のファイルパス"
Move_Folda = r"処理対象所属までのファイルパス" 

# subject
for Parent_Path in glob.glob(File_path + "\\*"):
   Shogo_Name_Folda =  Move_Folda  + "\\" + os.path.basename(Parent_Path)
   if os.path.exists(Shogo_Name_Folda) == False:
       os.mkdir(Shogo_Name_Folda)
   
   for Folda_Path in glob.glob(Parent_Path + "\\*"):
      Move_subject = []
      for File_path in glob.glob(Folda_Path + "\\*"):
         Move_subject.append(File_path)
      for item in Move_subject:

         if os.path.exists(Shogo_Name_Folda + "\\" +os.path.basename(item)) == False:     
            shutil.copyfile(item, Shogo_Name_Folda + "\\" +os.path.basename(item))
   # for item in Move_subject: 
      
   #    print(os.path.basename(item))