# Module
import os
import time
import pandas

startTime = time.time()

# Recuperation de la liste des fichiers du dossier
listFiles = os.listdir("aConvertir")
for fileId in range(0, len(listFiles)):
    listFiles[fileId] = listFiles[fileId].split('.')[0]

for fileName in listFiles:
    try:
        # Ouverture du fichier .txt a convertir
        file = open("aConvertir/"+fileName+".txt", 'r')
        assert fileName != "readme"
        
    except Exception:
        pass 
    
    else:
        # Lecture du fichier
        new_text = file.readlines()

        # Creation d'une liste contenant tout les mots dans le fichier et initialisation variables
        words = []
        line_break = 0

        # Ajout des mots du fichier dans la liste
        for x in range(0, len(new_text)):
            for word in new_text[x].split("\t"):
                words.append(word + ',')

        # Ecriture des mots dans le nouveau fichier .csv
        f = open("result/"+fileName+'.csv','w')

        for x in range(0, len(words)):
            if (line_break == 4):
                f.write('\n')
                f.write(words[x])
                line_break = 0
            elif line_break < 3:
                f.write(words[x])
                
            line_break += 1
        f.close()
        
df = pandas.read_csv('test.txt', delimiter=" ")
df.to_excel('output.xlsx', 'Sheet1', index=False)

# Confirmation de fin d'execution du script
print("########################")
print("# Conversion terminée. #")
print("########################")

print("# en {} secondes #".format(time.time()-startTime))