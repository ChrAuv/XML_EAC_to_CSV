from xml.etree import ElementTree
from csv import writer
from time import time
# J'importe ici l''ensemble de "os"
import os
# https://docs.python.org/fr/3/library/pathlib.html
from pathlib import Path


# script de transformation des fichiers d'un répertoire EAC
# vers un fichier CSV
# Christophe Auvray, Archives nationales du monde du travail


def fill_cell(XPath):
    """Cette fonction crée une cellule à partir de la valeur d'un élément EAC.
    Si l'élément est trouvé (try), la cellule est créée avec sa valeur.
    Si l'élément n'est pas trouvé (except), la cellule est créée vide."""
    try:
        element = root.find(XPath)
        # Permet de prendre en compte les balises <p>, <emph>, <list>, etc.
        element = ''.join(element.itertext())
        # Supprimer les espaces en trop avec les méthodes split() et join() 
        element = ' '.join(element.split())
    except AttributeError as e:
        element = ''
    return element


print("Ce script va transformer l'ensemble d'un répertoire de fichiers en format XML-EAC "
      "en un seul fichier CSV qui reprendra l'essentiel des données.\n"
      "Les fichiers doivent se trouver dans le répertoire ./EAC, au même niveau "
      "que le script.")

# Définir le répertoire où se situent les fichiers XML
xml_directory = './EAC'

input("Appuyer sur Entrée pour continuer")

# Créer le fichier CSV
with open('output_EAC.csv', 'w', encoding="utf-8", newline="") as f:
    write_file = writer(f, delimiter=';')

    # Définir les en-têtes de colonnes
    headers = ["Fichier XML", "Identifiant de la notice", "Code de l’organisme", "Type d'entité", "Forme du nom - Partie", "Forme du nom - Forme alternative",
               "Dates d’existence - Date de début", "Dates d’existence - Date de fin", "Lieu - Adresse(1)", "Lieu - Adresse(2)", "Lieu - Adresse(3)", "Lieu - Adresse(4)",
               "Statut juridique - Terme", "Statut juridique - Date de début", "Statut juridique - Date de fin", "Statut juridique - Note descriptive", "Fonctions - Terme",
               "Fonctions - Note descriptive", "Organisation interne ou généalogie", "Biographie ou histoire"]

    # Ecrire les en-têtes de colonne
    write_file.writerow(headers)

    # Dire que l'on veut travailler sur chaque fichier XML du répertoire
    # Si j'utilise print(xml_files_list) je constate que cette liste fonctionne
    xml_files_list = list(map(str,Path(xml_directory).glob('**/*.xml')))

    # Heure de début
    time_begin = time()

    # Parser chaque fichier XML (boucle)
    for xml_file in xml_files_list:
        # Le nom de chaque fichier traité va s'afficher
        print(xml_file)
        tree = ElementTree.parse(xml_file)
        # définit la racine (contexte XPath)
        root = tree.getroot()
       
        # Extraire les données
        # voir https://stackoverflow.com/questions/46959107/python-parsing-xml-to-csv-with-missing-elements

        # vigilance : l'espace de nom xmlns. Quand j'ajoute namespace = "{urn:isbn:1-931666-33-4}" ainsi que
        # entityType = root.find('.//{0}cpfDescription/{0}identity/{0}entityType'.format(namespace)).text, ça marche !
        # Sinon, ne fonctionnera pas
        # ici, espace de nom de l'EAC indiqué
        namespace = "{urn:isbn:1-931666-33-4}"

        # Nom du fichier XML
        # os.path.basename() - qui nécessite import os - permet d'avoir uniquement le nom du fichier avec l'extension,
        # sans son chemin
        name = os.path.basename(xml_file)
        
        # Identifiant de la notice
        recordId = fill_cell('.//{0}control/{0}recordId'.format(namespace))

        # Code de l’organisme
        agencyCode = fill_cell('.//{0}control/{0}maintenanceAgency/{0}agencyCode'.format(namespace))

        # Type d'entité
        entityType = fill_cell('.//{0}cpfDescription/{0}identity/{0}entityType'.format(namespace))

        # Forme du nom - Partie
        part = fill_cell('.//{0}cpfDescription/{0}identity/{0}nameEntry/{0}part'.format(namespace))

        # Forme du nom - Forme alternative
        alternativeForm = fill_cell('.//{0}cpfDescription/{0}identity/{0}nameEntry/{0}alternativeForm'.format(namespace))

        # Dates d’existence - Date de début
        existDatesFromDate = fill_cell('.//{0}cpfDescription/{0}description/{0}existDates/{0}dateRange/{0}fromDate'.format(namespace))

        # Dates d’existence - Date de fin
        existDatesToDate = fill_cell('.//{0}cpfDescription/{0}description/{0}existDates/{0}dateRange/{0}toDate'.format(namespace))

        # Lieu - Adresse (1)
        addressLine1 = fill_cell('.//{0}cpfDescription/{0}description/{0}place/{0}address/{0}addressLine[1]'.format(namespace))

        # Lieu - Adresse (2)
        addressLine2 = fill_cell('.//{0}cpfDescription/{0}description/{0}place/{0}address/{0}addressLine[2]'.format(namespace))

        # Lieu - Adresse (3)
        addressLine3 = fill_cell('.//{0}cpfDescription/{0}description/{0}place/{0}address/{0}addressLine[3]'.format(namespace))
        
        # Lieu - Adresse (4)
        addressLine4 = fill_cell('.//{0}cpfDescription/{0}description/{0}place/{0}address/{0}addressLine[4]'.format(namespace))
        
        # Statut juridique - Terme
        legalStatusterm = fill_cell('.//{0}cpfDescription/{0}description/{0}legalStatuses/{0}legalStatus/{0}term'.format(namespace))
        
        # Statut juridique - Date de début
        legalStatusFromDate = fill_cell('.//{0}cpfDescription/{0}description/{0}legalStatuses/{0}legalStatus/{0}dateRange/{0}fromDate'.format(namespace))

        # Statut juridique - Date de fin
        legalStatusToDate = fill_cell('.//{0}cpfDescription/{0}description/{0}legalStatuses/{0}legalStatus/{0}dateRange/{0}toDate'.format(namespace))

        # Statut juridique - Note descriptive
        legalStatusNote = fill_cell('.//{0}cpfDescription/{0}description/{0}legalStatuses/{0}legalStatus/{0}descriptiveNote'.format(namespace))
        
        # Fonctions - Terme
        functionsTerm = fill_cell('.//{0}cpfDescription/{0}description/{0}functions/{0}function/{0}term'.format(namespace))
   
        # Fonctions - Note descriptive
        functionsNote = fill_cell('.//{0}cpfDescription/{0}description/{0}functions/{0}function/{0}descriptiveNote'.format(namespace))
  
        # Organisation interne ou généalogie
        structureOrGenealogy = fill_cell('.//{0}cpfDescription/{0}description/{0}structureOrGenealogy'.format(namespace))

        # Biographie ou histoire
        biogHist = fill_cell('.//{0}cpfDescription/{0}description/{0}biogHist'.format(namespace))
 
        # écriture de la ligne dans le CSV
        csv_line = [name, recordId, agencyCode, entityType, part, alternativeForm, existDatesFromDate, existDatesToDate, addressLine1, addressLine2, addressLine3, addressLine4,
                    legalStatusterm, legalStatusFromDate, legalStatusToDate, legalStatusNote, functionsTerm, functionsNote, structureOrGenealogy, biogHist]

        # ajouter une nouvelle ligne au fichier CSV avec les données
        write_file.writerow(csv_line)

    # Heure de fin
    time_end = time()

    # Temps total
    total_time = time_end - time_begin


print(f"Un total de {len(xml_files_list)} fichiers à été parsé en {total_time} "
      f"secondes. Le fichier produit se trouve dans {os.getcwd()}. "
      f"Il comporte {len(headers)} colonnes.")

                
