#This script help the redaction of the Seveso / ICPE report for all products in stock + laboratories in lyon.
#It works by processing 3 excel sheets
#   -VC+ extraction: column used : ["PRODUCT_NAME", "STR_ID", "STOCK_AMOUNT","STOCK_AMOUNT_UNIT", "CAS_NO", "STORAGE_CLASS", "SAFETY_PHRASES", "PHYSICAL_STATE"
#   -A table containing all named_substances, their classification (a, b, c), and all cas numbers
#   -A table containing all thresholds (seveso and ICPE) detected in the VC+ extraction, and on wich classification they apply

# Import all packages 
import numpy as np
import pandas as pd
import re

#Charger les datas STOCK + Laboratoire Lyon et changer les NaN par ""
pd.set_option('display.max_columns', None)
df = pd.read_excel("Data/Axel LYO Vials.xlsx", dtype={'STR_ID' : str,
                                                      'VIAL_BARCODE' : str,
                                                      'LAB_ID': str,
                                                      'LOCATION_ID': str})
df = df.fillna("")

#Conversion des quantités en Tonne (pour les liquides, une densité de 1 est appliquée)

#Renommer les colonnes contenant les masses originales pour les conserve (verification) et envoyer la transformation de la convertion en tonnes en fin de dataframe
df.rename(columns = {"STOCK_AMOUNT" : "STOCK_AMOUNT_vc",
                     "STOCK_AMOUNT_UNIT" : "STOCK_AMOUNT_UNIT_vc"}, inplace=True)


    #PCE pour correspondre à l'unité et à la capacité du vial
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == "PCE"), ["STOCK_AMOUNT_vc"]] = df['VIAL_CAPACITY']
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == "PCE"), ["STOCK_AMOUNT_UNIT_vc"]] = df['VIAL_CAPACITY_UNIT']
    #g and ml in tonne
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'g') | (df['STOCK_AMOUNT_UNIT_vc'] == 'ml'), ['STOCK_AMOUNT']] = df['STOCK_AMOUNT_vc'] / 1000000
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'g') | (df['STOCK_AMOUNT_UNIT_vc'] == 'ml'), ['STOCK_AMOUNT_UNIT']] = 'ton'
    #kg and l in tonne
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'kg') | (df['STOCK_AMOUNT_UNIT_vc'] == 'l'), ['STOCK_AMOUNT']] = df['STOCK_AMOUNT_vc']/1000
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'kg') | (df['STOCK_AMOUNT_UNIT_vc'] == 'l'), ['STOCK_AMOUNT_UNIT']] = 'ton'
    #mg in tonne
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'mg'), ['STOCK_AMOUNT']] = df['STOCK_AMOUNT_vc']/1000000000
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'mg'), ['STOCK_AMOUNT_UNIT']] = 'ton'
    #stone en tonne
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'st'), ['STOCK_AMOUNT']] = df['STOCK_AMOUNT_vc']/157.5
df.loc[(df['STOCK_AMOUNT_UNIT_vc'] == 'st'), ['STOCK_AMOUNT_UNIT']] = 'ton'

#Charger le tableau des références des substances nommément désignées et nettoyage des données
df_ref_name = pd.read_excel("Data/named_substances.xlsx")
df_ref_name = df_ref_name.fillna('0')  

df_ref_name['phrase_H'] = df_ref_name['phrase_H'].str.split(r"\n") 
df_ref_name['CAS'] = df_ref_name['CAS'].str.split(', ')
df_ref_name = df_ref_name.explode('CAS')

#Création d'une colonne "regle cumul"
df_ref_name['regle cumul'] = ""
df_ref_name.loc[df_ref_name['a'] == "x", ['regle cumul']] += 'a, '
df_ref_name.loc[df_ref_name['b'] == "x", ['regle cumul']] += 'b, '
df_ref_name.loc[df_ref_name['c'] == "x", ['regle cumul']] += 'c, '


df_ref_name = df_ref_name.rename(columns={"CAS" : "CAS_NO"}) #Renaming CAS column so it matches the database

#////////////////////////////////////////////////////////////////////////////////////// Classification des produits Nommément désignés avec le df_name_ref

# Rajout de la colonne 'Nommement designee" et attribution des valeurs "Oui" ou "Non"
df.loc[df['CAS_NO'].isin(df_ref_name['CAS_NO']), ['Nommement designee ?']] = 'Oui'
df.loc[~(df['Nommement designee ?'] == 'Oui'), ['Nommement designee ?']] = 'Non'

#Rajout de la colonne "regle cumul" et remplissage pour produits nommément designés avec rubrique
cumul = df_ref_name.loc[:, ["CAS_NO", "regle cumul", "rubrique"]]
df = df.merge(cumul, how="left", on="CAS_NO")

#Nettoyage des colonnes "regle cumul" et "rubrique"
df[["regle cumul", "rubrique"]] = df[["regle cumul", "rubrique"]].fillna("")

#////////////////////////////////////////////////////////////////////////////////////// Classification des produits NON nomméments désignés
#Catégorie phrases H pour l'attribution de la règle de cumul à la liste de produits
H_toxiques = ['H300', 'H310', 'H330', 'H331', 'H301', 'H370']
H_physiques = ['H200', 'H201','H202','H203','H204','H205', 'H220', 'H221', 'H222', 'H223', 'H224', 'H225', 'H226', 'H240', 'H241', 'H242', 'H250', 'H270', 'H271', 'H272', 'H270']
H_environnements = ['H400', 'H410', 'H411'] 

#Completion de la colonne "rubrique" et "regle cumul" pour tous les autres composés
#Pas de differenciation solide / liquide / gas, la catégorie contenant le seuil le plus bas est systématiquement appliquée pour les calculs


    #Classification Toxique
    
#Toxicite aigu cat 1
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['STORAGE_CLASS'] == "6.1A")
         |(df['STORAGE_CLASS'] == "6.1B"))
       &((df['SAFETY_PHRASES'].str.contains("H300")) 
        |(df['SAFETY_PHRASES'].str.contains("H310")) 
        |(df['SAFETY_PHRASES'].str.contains("H330"))), ['rubrique']] += "4110, "

#Si "storage class" n'est pas détaillé => ajout en 4110
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['STORAGE_CLASS'] != "6.1A")
         &(df['STORAGE_CLASS'] != "6.1B")
         &(df['STORAGE_CLASS'] != "6.1C")
         &(df['STORAGE_CLASS'] != "6.1D"))
       &((df['SAFETY_PHRASES'].str.contains("H300")) 
        |(df['SAFETY_PHRASES'].str.contains("H310")) 
        |(df['SAFETY_PHRASES'].str.contains("H330"))), ['rubrique']] += "4110, "

#Toxicite aigu cat 2
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['STORAGE_CLASS'] == "6.1C")
         |(df['STORAGE_CLASS'] == "6.1D"))
       &((df['SAFETY_PHRASES'].str.contains("H300")) 
        |(df['SAFETY_PHRASES'].str.contains("H310")) 
        |(df['SAFETY_PHRASES'].str.contains("H330"))), ['rubrique']] += "4120, "


#Toxicite aigu cat 3 par inhalation
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("H331")), ['rubrique']] += "4130, "


#Toxicite aigu cat 3 voie d'exposition orale
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("H301")), ['rubrique']] += "4140, "

#Toxicite specifique
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("H370")), ['rubrique']] += "4150, "

#Attribution de la règle de cumul a
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("|".join(H_toxiques))), ['regle cumul']] += "a, "

#Classification dangers physiques
    
#Gaz inflammables et extrêmement inflammables
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['SAFETY_PHRASES'].str.contains("H220")) 
        |(df['SAFETY_PHRASES'].str.contains("H221"))) , ['rubrique']] += "4310, "

#Aerosols inflammables / Categorie 4320 à différencier manuellement de 4321 si dépassement du seuil
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['SAFETY_PHRASES'].str.contains("H222")) 
        |(df['SAFETY_PHRASES'].str.contains("H223"))) , ['rubrique']] += "4320, "

#Liquides et vapeurs inflammables cat 1
df.loc[(df['Nommement designee ?'] == "Non") 
       & (df['SAFETY_PHRASES'].str.contains("H224")), ['rubrique']] += "4330, "

#Liquides inflammables cat 2 & 3
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['SAFETY_PHRASES'].str.contains("H225")) 
        |(df['SAFETY_PHRASES'].str.contains("H226"))), ['rubrique']] += "4331, "

#Substances et mélanges auto-reactifs type A & B, hors peroxydes                           
df.loc[(df['Nommement designee ?'] == "Non") 
       &(~df['PRODUCT_NAME'].str.contains("perox", case=False, na=False))
       &((df['SAFETY_PHRASES'].str.contains("H240")) 
        |(df['SAFETY_PHRASES'].str.contains("H241"))), ['rubrique']] += "4410, "

#Substances et mélanges auto-reactifs type C, D, E, F hors peroxydes                       
df.loc[(df['Nommement designee ?'] == "Non") 
       &(~df['PRODUCT_NAME'].str.contains("perox", case=False, na=False))
       &(df['SAFETY_PHRASES'].str.contains("H242")), ['rubrique']] += "4411, "


#Peroxydes organiques type A & B                                                           
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['PRODUCT_NAME'].str.contains("perox", case=False, na=False))
       &((df['SAFETY_PHRASES'].str.contains("H240")) 
        |(df['SAFETY_PHRASES'].str.contains("H241"))), ['rubrique']] += "4420, "

#Peroxydes organiques type C, D, E, F (pas de differentiation typ C,D et type E,F) rub. 4221 et 4222 combinées      
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['PRODUCT_NAME'].str.contains("perox", case=False, na=False))
       &(df['SAFETY_PHRASES'].str.contains("H242")), ['rubrique']] += "4421, "

#Liquides pyrophoriques catégorie 1
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['PHYSICAL_STATE'] == "liquid")
       &(df['SAFETY_PHRASES'].str.contains("H250")), ['rubrique']] += "4430, "

#Matières solides pyrophoriques catégorie 1
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['PHYSICAL_STATE'] == "solid")
       &(df['SAFETY_PHRASES'].str.contains("H250")), ['rubrique']] += "4431, "

#Liquides comburants catégorie 2, 3
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['PHYSICAL_STATE'] == "liquid")
       &((df['SAFETY_PHRASES'].str.contains("H271"))
        |(df['SAFETY_PHRASES'].str.contains("H272"))), ['rubrique']] += "4440, "

#Solides comburants catégorie 2, 3
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['PHYSICAL_STATE'] == "solid")
       &((df['SAFETY_PHRASES'].str.contains("H271"))
        |(df['SAFETY_PHRASES'].str.contains("H272"))), ['rubrique']] += "4441, "

#Gaz comburants catégorie 1
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("H270")), ['rubrique']] += "4442, "

#Recherche explosifs => considérés non emballé conformément aux dispositions de transport => detection en rubrique 4240.1 seulement
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['SAFETY_PHRASES'].str.contains("H200"))
         |(df['SAFETY_PHRASES'].str.contains("H201"))
         |(df['SAFETY_PHRASES'].str.contains("H202")) 
         |(df['SAFETY_PHRASES'].str.contains("H203")) 
         |(df['SAFETY_PHRASES'].str.contains("H204"))  
         |(df['SAFETY_PHRASES'].str.contains("H205"))), ['rubrique']] += "4240.1, "

#Substances et mélanges qui, au contact de l'eau, dégagent des gaz inflammables, cat1
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("H260")), ['rubrique']] += "4620, "

#Réagit violemment au contact de l'eau
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("EUH014")), ['rubrique']] += "4610, "

#Au contact de l'eau, dégage des gaz toxiques
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("EUH029")), ['rubrique']] += "4630, "

#Attribution de la règle de cumul b
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("|".join(H_physiques))), ['regle cumul']] += "b, "

#Classification dangers environnement
    
#Dangers pour le milieu aquatique - Danger aigu et chronique, cat 1
df.loc[(df['Nommement designee ?'] == "Non") 
       &((df['SAFETY_PHRASES'].str.contains("H400")) 
        |(df['SAFETY_PHRASES'].str.contains("H410"))) , ['rubrique']] += "4510, "

#Dangers pour le milieu aquatique - Danger chronique, cat 2
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("H411")), ['rubrique']] += "4511, "

#Attribution de la règle de cumul c
df.loc[(df['Nommement designee ?'] == "Non") 
       &(df['SAFETY_PHRASES'].str.contains("|".join(H_environnements))), ['regle cumul']] += "c, "

# /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
# Grouper les produits par "IDs" et spécification de l'aggrégation de chaque colomne
df_id = df.groupby('STR_ID').agg({'PRODUCT_NAME' : lambda x: ', '.join(x.unique()), 
                                  'CAS_NO' : lambda x: ', '.join(x.unique()),
                                  'STOCK_AMOUNT' : lambda x: x.sum().round(7), 
                                  'STOCK_AMOUNT_UNIT' : lambda x: ', '.join(x.unique()), 
                                  'CAS_NO' : lambda x: ', '.join(x.unique()),
                                  'SAFETY_PHRASES' : lambda x: ', '.join(x.unique()),
                                  'Nommement designee ?': 'unique',
                                  'rubrique' : lambda x: ', '.join(x.unique()),
                                  'regle cumul':  lambda x: ', '.join(x.unique())
                                  }).sort_values('STOCK_AMOUNT', ascending=False).reset_index()

#Nettoyage de la colonne CAS_NO
df_id['CAS_NO'] = df_id['CAS_NO'].str.strip()
df_id['CAS_NO'] = df_id['CAS_NO'].str.strip(',')
df_id['CAS_NO'] = df_id['CAS_NO'].str.strip()

#Nettoyage de la colonne safety phrase
def parse(chain):
    processed = ''
    chain_list = list(set(chain.split(', ')))
    chain_list.sort()
    for i in chain_list:
        if i !='':
            processed += (i + ", ")
    return processed  

df_id['SAFETY_PHRASES'] = df_id['SAFETY_PHRASES'].apply(lambda x : parse(x))

#Nettoyage de la colonne Nommément designé 
df_id['Nommement designee ?'] = df_id['Nommement designee ?'].str.join(', ')
df_id.loc[df_id['Nommement designee ?'].str.contains("Oui"), ['Nommement designee ?']] = "Oui"
df_id.loc[df_id['Nommement designee ?'].str.contains("Non"), ['Nommement designee ?']] = "Non"

#Nettoyage de la colomne "rubrique" et supression des doublons 
df_id['rubrique'] = df_id['rubrique'].apply(lambda x : parse(x))

#Nettoyage colonne regle cumul
df_id['regle cumul2'] = df_id['regle cumul'].copy()
df_id['regle cumul'] = ''
df_id.loc[(df_id['regle cumul2'].str.contains("a")), ['regle cumul']] += 'a, '
df_id.loc[(df_id['regle cumul2'].str.contains("b")), ['regle cumul']] += 'b, '
df_id.loc[(df_id['regle cumul2'].str.contains("c")), ['regle cumul']] += 'c, '
df_id.drop('regle cumul2', axis=1, inplace=True)

#chargement de la table pour seuil seveso et ICPE et préparation colonne 'rubrique'
def parse_seuils(x):
    x = str(x)
    if x[-1] == '0':
        x = x[:-2]
    return x

df_seuils = pd.read_excel("Data/Seuils_seveso_ICPE.xlsx")
df_seuils['rubrique'] = df_seuils['rubrique'].astype(str)
df_seuils['rubrique'] = df_seuils['rubrique'].apply(lambda x: parse_seuils(x))

#Séparer toutes les lignes par rubriques
df_id['rubrique'] = df_id['rubrique'].str.split(', ')
df_id = df_id.explode('rubrique')

#Renseignement des seuils bas seveso, ICPE et type de danger a, b ou c ou groupes de molécules
df_id = df_id.merge(df_seuils, how="left", on="rubrique")

#Creation des colonnes seuil a, b et c et remplissage du seuil correspondant à chacune des catégories pour les non Nommement designee
df_id.loc[(df_id['cumul'] == "a")&(df_id['Nommement designee ?'] == 'Non'), ['seuil cumul a']] = df_id['seuil seveso bas']
df_id.loc[(df_id['cumul'] == "b")&(df_id['Nommement designee ?'] == 'Non'), ['seuil cumul b']] = df_id['seuil seveso bas']
df_id.loc[(df_id['cumul'] == "c")&(df_id['Nommement designee ?'] == 'Non'), ['seuil cumul c']] = df_id['seuil seveso bas']


#Regrouper les données en prenant les seuils minimums et en considérant les règles de cumuls applicables à chacune des catégories
df_id = df_id.groupby('STR_ID').agg({'PRODUCT_NAME' : 'first', 
                                  'CAS_NO' : 'first',
                                  'STOCK_AMOUNT' : 'first', 
                                  'STOCK_AMOUNT_UNIT' : 'first', 
                                  'CAS_NO' : 'first',
                                  'SAFETY_PHRASES' : 'first',
                                  'Nommement designee ?': 'first',
                                  'rubrique' : lambda x: ', '.join(x),
                                  'regle cumul':  'first',
                                  'seuil seveso bas': 'min', 
                                  'seuil ICPE bas': 'min',
                                  'seuil cumul a': 'min',
                                  'seuil cumul b':'min',
                                  'seuil cumul c': 'min'
                                  }).sort_values('STOCK_AMOUNT', ascending=False).reset_index()

#Ajout des seuils minimum de cumuls aux produits nommément désignés
df_id.loc[(df_id['Nommement designee ?'] == 'Oui')
          &(df_id['regle cumul'].str.contains('a')), ['seuil cumul a']] = df_id['seuil seveso bas']

df_id.loc[(df_id['Nommement designee ?'] == 'Oui')
          &(df_id['regle cumul'].str.contains('b')), ['seuil cumul b']] = df_id['seuil seveso bas']

df_id.loc[(df_id['Nommement designee ?'] == 'Oui')
          &(df_id['regle cumul'].str.contains('c')), ['seuil cumul c']] = df_id['seuil seveso bas']

#Calcul des fractions pour chaque produit
df_id['fraction a'] = (df_id['STOCK_AMOUNT'] / df_id['seuil cumul a'])
df_id['fraction b'] = (df_id['STOCK_AMOUNT'] / df_id['seuil cumul b'])
df_id['fraction c'] = (df_id['STOCK_AMOUNT'] / df_id['seuil cumul c'])

#Vérification Dépassement seuil direct Seveso Bas et ICPE Bas
df_id['Dépassement direct seuil seveso'] = df_id['STOCK_AMOUNT'] > df_id['seuil seveso bas']
df_id['Dépassement direct seuil ICPE'] = df_id['STOCK_AMOUNT'] > df_id['seuil ICPE bas']

#Calcul des sommes par type de danger :
df_id = df_id.fillna(0)
a = sum(df_id['fraction a'])
b = sum(df_id['fraction b'])
c = sum(df_id['fraction c'])
print("Somme a = " + str(a))
print("Somme b = " + str(b))
print("Somme c = " + str(c))

#Création d'un Datafram Resultat
df_resultat = df_id.loc[(df_id['Dépassement direct seuil seveso'] == True)
          |df_id['Dépassement direct seuil ICPE'] == True] 

#Write results on excel document
with pd.ExcelWriter("Seveso_ICPE_Lyon.xlsx") as writer:
    df.to_excel(writer, sheet_name="Inventaire Detaillé")
    df_id.to_excel(writer, sheet_name="Inventaire groupé par produit")
    df_resultat.to_excel(writer, sheet_name="Depassement Seveso ou ICPE")