# Module
import os
import time
import pandas

startTime = time.time()

choix = int(input("Choix de conversion (0->CSV ou 1->EXCEL)\n"))

SOURCE_ENCODING = "utf-16 LE"
TARGET_ENCODING = "utf-8"

# Recuperation de la liste des fichiers du dossier
listFiles = os.listdir("aConvertir")
for fileId in range(0, len(listFiles)):
    listFiles[fileId] = listFiles[fileId].split('.')[0]

if choix==0 or choix==1 :
    for fileName in listFiles:
        try:
            # Ouverture du fichier .txt a convertir
            file = open("aConvertir/"+fileName+".txt", 'r')
            assert fileName != "readme" 
            
        except Exception:
            pass 
        
        else:
            if choix == 0 :
                # Lecture du fichier
                new_text = file.readlines()

                # Creation d'une liste contenant tout les mots dans le fichier et initialisation variables
                words = []
                line_break = 0

                # Ajout des mots du fichier dans la liste
                for x in range(0, len(new_text)):
                    for word in new_text[x].split(";"):
                        words.append(word + ',')

                # Ecriture des mots dans le nouveau fichier .csv
                f = open("result/"+fileName+".csv",'w')

                for x in range(0, len(words)):
                    if (line_break == 4):
                        f.write('\n')
                        f.write(words[x])
                        line_break = 0 
                    elif line_break < 3:
                        f.write(words[x])
                        
                    line_break += 1
                f.close()
            elif choix == 1 :
                fileUtf16 = "aConvertir/"+fileName+".txt"
                fileUtf8 = "aConvertir/"+fileName+"_utf8.txt"
                
                with open(fileUtf16, "r", encoding=SOURCE_ENCODING) as fich :
                    content = fich.read()
                with open(fileUtf8, "w", encoding=TARGET_ENCODING) as fich :
                    fich.write(content)
                            
                df = pandas.read_csv("aConvertir/"+fileName+"_utf8.txt", sep=";")
                df.to_excel("result/"+fileName+".xlsx", 'Sheet1', index=False)
                
                print("Conversion du fichier {} : OK".format(fileName+".txt"))
else:
    print("Choix de conversion incorrecte.")

# Confirmation de fin d'execution du script
print("########################")
print("# Conversion terminée. #")
print("########################")

print("# en {} secondes #".format(time.time()-startTime))