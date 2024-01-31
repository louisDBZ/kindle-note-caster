### import

try: 
    import pandas as pd
    import docx
    import sys
except ImportError as e:
    print("please install the required dependancies. For that, run the requirement 'pip install - requirements.txt'. If needed, create & activate a virtual env with 'python -m venv venv'")
    

### Nom des chemins

if len(sys.argv) < 2:
	print("Usage: python kindle-note-caster.py Chemin-URI-csv")
	sys.exit(1)

i_chemin_du_csv= sys.argv[1]
o_chemin_du_word= i_chemin_du_csv[:-3] + "docx"

### fonctions

def ajouter_surlignement(surlignement):
    global o_texte
    if (str(surlignement) not in List_mot_a_enlever):
        o_texte = o_texte +'\n\n'+ str(surlignement)

def ajouter_note(note):
    global o_texte
    if (str(note) not in List_mot_a_enlever):
        o_texte = o_texte +' => '+ str(note)

### script

df=pd.read_csv(i_chemin_du_csv, sep=',') #on_bad_lines{‘error’, ‘warn’, ‘skip’}

o_texte = ""
List_mot_a_enlever=['nan','Annotation']

df.columns = ['Type','Emplacement','unknown','texte']

for i in df.index: 
    
    if df['Type'][i]=='Surlignement (Jaune)':
        ajouter_surlignement(df['texte'][i])
        
    elif df['Type'][i]=='Note':
        ajouter_note(df['texte'][i])

mydoc = docx.Document()
mydoc.add_paragraph(o_texte)
mydoc.save(o_chemin_du_word)

print("Word file is named: "+ o_chemin_du_word)