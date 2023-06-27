import os, glob, csv, shutil, math, logging, sys, pyodbc, matplotlib.pyplot as plt
from datetime import datetime
import traceback
import numpy as np
import openpyxl
import pdb
import simple_colors
from matplotlib.backends.backend_pdf import PdfPages
from dateutil.parser import parse
from time import *

TEST_MODE = False
if "testEnvironement" in os.getcwd() : TEST_MODE = True

#Récupération du path
current_directory = os.getcwd()

#récupération des fichiers .csv
search_filecsv = glob.glob('*.csv')
search_filecsv.sort()
search_fileuhh = glob.glob('*.uhh')
search_fileuhh.sort()
# Obtenir la date et l'heure actuelles
maintenant = datetime.now()

# Extraire le jour, le mois et l'année de l'objet `maintenant`
jour = maintenant.day
mois = maintenant.month
annee = maintenant.year

#création du dossier automoved
destination_directory = os.path.join(current_directory, "automoved" + str(annee))
os.makedirs(destination_directory, exist_ok=True)

#initialisation du logger avec logging (les logs vont etre enregistre dans le document nomme ci-dessous apres "filename")
logger = logging.getLogger(__name__)
logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', filename='Sebastien.log', filemode='a', level=logging.INFO)
logger.info('Process start')

#Connection a la database ODBC de HeliosII
conn_string = "driver=oracle; dbq=HELIOSII; uid=ETAT; pwd=ETAT" if TEST_MODE else "driver=oracle; dbq=HELIOSII; uid=ADMIN; pwd=ADMIN"
cnx = pyodbc.connect(conn_string)
logger.info("Connected using : " + conn_string)

#Creation de 4 curseur pour y enregistrer les résultats de requêtes SQL
csr1_CD_OFS, csr2_DONNEES, csr3, csr4, csr5 = [cnx.cursor() for _ in range(5)] # definit tout en tant que nouveau curseur

csr4.execute("SELECT DOCUMENT FROM DOC_TECH_DETAIL WHERE VALIDE =" + str(1) + " AND CD_DOC_TECH_ENTETE = " + str(137422))
donnee_prog = csr4.fetchone()
with open(current_directory + "\\" + "IMP650 (Critères de validation de traitementthermique).xlsx", "wb") as ff:
	for i in range(len(donnee_prog)):
		ff.write(donnee_prog[i])
ff.close()

#permet de ressortir la couleure d'un "OK"(Vert) ou "NON OK"(Rouge)
def verif_color(n):
	if n == "OK" :
		return("green")
	elif n == "NON OK" :
		return("red")
	else :
		return("black")

#Transformation d'une Donnée en heure minute vers une Donnée en minute		
def hm_in_minute(hm):
	minute = hm.split(":")
	minute[1] = int(minute[0])*60 + int(minute[1])
	return(minute[1])

#Permet de ressortir en fonction de la validité de la donnée par rapport a ce qui est éxiger un OK ou un NON OK ou un "-" si il n'y a pas d'éxigence du client	
def comparaison_temps(result,temps_max,temps_min) :
	if  temps_min != "-" and hm_in_minute(result) <  temps_min :
		conform_temps_maintien_min = "NON OK"
	elif   temps_min != "-" and hm_in_minute(result) >  temps_min :
		conform_temps_maintien_min = "OK"
	else :
		conform_temps_maintien_min = "-"
		
	if temps_max != "-" and hm_in_minute(result) > temps_max :
		conform_temps_maintien_max = "NON OK"
	elif  temps_max != "-" and hm_in_minute(result) < temps_max :
		conform_temps_maintien_max = "OK"
	else :
		conform_temps_maintien_max = "-"
		
	if conform_temps_maintien_max == "OK" and  conform_regulation_min == "OK":
		conform_temps_maintien = "OK"
	elif  temps_min == "-" and temps_max == "OK" or temps_max == "-" and conform_temps_maintien_min == "OK" :
		conform_temps_maintien = "OK"
	elif conform_temps_maintien_max == "NON OK" or conform_temps_maintien_min == "NON OK" :
		conform_temps_maintien = "NON OK"
	else :
		conform_temps_maintien = "-"
	return(conform_temps_maintien)

for file in search_fileuhh :
		source_path = os.path.join(current_directory, file)
		destination_path = os.path.join(destination_directory, file)
		shutil.move(source_path,destination_path)

if search_filecsv == "" :
	print("Il n'ya aucun fichier dans le dossier")
	sleep(10)
	
fin_programme = 0
reponse_mode = False
while reponse_mode == False :
	mode_auto = input("Veuillez saisir si vous souhaitez faire la manipulation en mode auto : 0 ou en mode manuel : 1 :    ")
	if mode_auto in ["0", "1"] :
		reponse_mode = True
		mode_auto = int(mode_auto)
	else :
		print("Attention le mode " + mode_auto + " n'éxiste pas")
#Parcours des fichiers Groupe-1 et Group-2 et du fichier IMP650 dans le répertoire	
		
for filepath in search_filecsv:
	try :
		if "--" not in filepath :
			source_path = os.path.join(current_directory, filepath)
			destination_path = os.path.join(destination_directory, filepath)
			shutil.move(source_path, destination_path)
		move_after_verif = False
		# Vérification de l'éxistence du fichier Groupe-1 et IMP650
		if filepath.startswith("Group-1") and "E-----------------" not in filepath and "--" in filepath:
			groupe_1 = filepath
			groupe_2 = groupe_1.replace("Group-1","Group-2")
			
			if groupe_2 not in search_filecsv:
				raise FileNotFoundError("Il n'y a aucun fichier groupe-2 correspondant au fichier " + groupe_1 + "\n\nVeuillez vérifiez si il n'y a pas de faute dans le nom des fichiers\n")
			else:
				numero_programme = []
				consigne_grp2 = []
				move_after_verif = True
				
				#ouverture des fichiers
				groupe_1_open = open(groupe_1, newline = '')
				groupe_2_open = open(groupe_2, newline = '')
				wb = openpyxl.load_workbook(current_directory + "\\" + "IMP650 (Critères de validation de traitementthermique).xlsx")
				sheet_IMP650 = wb['IMP650']
				cr_groupe2 = csv.reader( groupe_2_open,dialect='excel',delimiter=';')
				cr_groupe1 = csv.reader(groupe_1_open,dialect='excel' ,delimiter=';')
				
				
				#Vérification si le numéro de programme a changé au cours du processus
				ligne_actuelle = 1
				for row in cr_groupe2:
					if ligne_actuelle >= 11:
						if row[1] == '' :
							numero_programme.append(row[5])
							consigne_grp2.append(row[4])
					ligne_actuelle += 1
					
				dernier_consigne = float(consigne_grp2[-1].replace(",","."))	
				changement_programme = False
				if len(numero_programme) >= 1:
					dernier_programme = numero_programme[-1]
					
					for donnee in numero_programme:
						if donnee != dernier_programme:
							print("ATTENTION IL Y A EU UN CHANGEMENT DE PROGRAMME AU COURS DU PROCESSUS")
							logger.info("ATTENTION IL Y A EU UN CHANGEMENT DE PROGRAMME AU COURS DU PROCESSUS")
							changement_programme = True
							break
				
				#récupération des of dans les titres
				heliosinfo = groupe_1.split("~")
				if len(heliosinfo) > 1 :
						heliosinfo = heliosinfo[1] # recupere les OFs et le numero de fournee
						
				else:
					print("heliosinfo list is empty, demande_DONNESping this file")
					continue
				fournee = heliosinfo.split("--")[1] # separe le numero de fournee
				OFs = heliosinfo.split("--")[0].split("-") # recupere le ou les OFs dans une liste
				OFs.sort() # tri des OFs dans l'ordre croissant
				
				Key = fournee + " ".join(OFs) # n° de founee suivi de tout les OFs
				
				#Parcours de tout les OFs pour y récupérer les données et les exploiter
				for OF in OFs:
					#Récupération du CD_OFS
					csr1_CD_OFS.execute("SELECT CD_OFS FROM OFS WHERE ID_OFS = " + OF)
					cd_ofs = csr1_CD_OFS.fetchone()
					
					#Récupération des données client, matière et épaisseur
					csr2_DONNEES.execute("WITH EPAISSEUR_MATIERE AS ( SELECT o.CD_ARTICLE , CASE WHEN INSTR(REF_LIBELLE, 'Epaisseur') <> 0 THEN SUBSTR(REF_LIBELLE,INSTR(REF_LIBELLE, 'Epaisseur')+12,INSTR(REF_LIBELLE, '(MM)',INSTR(REF_LIBELLE, 'Epaisseur'))-INSTR(REF_LIBELLE, 'Epaisseur')-13) ELSE Null END Epaisseur, SUBSTR(REF_LIBELLE, 1, INSTR(REF_LIBELLE, ' - ') - 1) Matiere FROM OFS_BESOIN ob INNER JOIN OFS o ON o.CD_OFS = ob.CD_OFS AND TRIM(ID_OFS) = " + str(OF) + ") SELECT DISTINCT em.Epaisseur, em.Matiere, c.NOM FROM Client c INNER JOIN Article a ON a.CD_CLIENT = c.CD_CLIENT INNER JOIN EPAISSEUR_MATIERE em ON a.CD_ARTICLE = em.CD_ARTICLE ORDER BY c.NOM, em.Matiere, em.Epaisseur")
					DONNEES_csr2 = csr2_DONNEES.fetchall()
					verif_epaisseur = False
					reponse_of = False
					continue_boolean = False
					if mode_auto == 1 :
						while not reponse_of :
							ii = []
							#Demande a l'utilisateur quel OF veut-il exploiter pour le processus et vérification de la validité de l'input
							if len(DONNEES_csr2) > 0 :
								
								for i in range(0,len(DONNEES_csr2)):
									if i < 1 :
										print("\nS  : Pour skip l'OF \n")
										print("E  : Pour arrêter le programme \n")
									ii.append(i)
									print(i , " : ", DONNEES_csr2[i], "\n")
								demande_DONNEES = input("Il y'a plusieurs données récupérées pour l'OF "+ OF+ ", veuillez renseigner la donnée que vous souhaitez exploiter avec le numéro sur la gauche de celle-ci :  ") 
								if (demande_DONNEES.isdigit() and int(demande_DONNEES) not in ii) and demande_DONNEES not in ["S", "s", "E", "e"]:
									print("Le choix que vous avez entré n'est pas possible")
									reponse_of = False
								else:
									reponse_of = True
									if demande_DONNEES in ["S", "s"]:
										continue_boolean = True
										continue
									if demande_DONNEES in ["E", "e"]:
										exit()
									DONNEES = DONNEES_csr2[int(demande_DONNEES)]
							#on met un deuxième continue car cela permet de passer a un autre of dans la boucle for
							if continue_boolean :
								continue
						if continue_boolean :
							continue
					else :
						DONNEES = DONNEES_csr2[0]
					nom_pdf = "TTH_" + fournee +"_" + OF + ".pdf"
					with PdfPages(nom_pdf) as pdf:	# Create a PDF file that contains the curves	
						text = ''
						font = {'family' : 'calibri',	# Set default texte properties
								'weight' : 'normal',
								'size'   : 20}

						plt.rc('font', **font)
						#création de la courbe pour le fichier PDF

						fig,ax = plt.subplots()
						fig.set_figheight(17)
						fig.set_figwidth(23)	
						saisi_epaisseur = False
						saisi_client = False
						saisi_matiere = False
							
						
						
						#Demande a l'utilisateur les données manquante en fonction de ce qu'il ressors
						if DONNEES[0] is None or DONNEES[1] is None :
							csr5.execute("WITH ART_REF AS ( SELECT DISTINCT TRIM(a.ID_ARTICLE) id_article, ob.REF_LIBELLE FROM OFS_BESOIN ob INNER JOIN OFS o ON ob.CD_OFS = o.CD_OFS INNER JOIN ARTICLE a ON o.CD_ARTICLE = a.CD_ARTICLE WHERE ob.REF_LIBELLE is not null AND ob.CD_OFS IN (SELECT CD_OFS FROM STOCK_ARTI_EOF WHERE CD_STOCK_ARTI_EOF IN (SELECT CD_STOCK_ARTI_EOF FROM STOCK_ARTI_SOF WHERE CD_OFS = (SELECT CD_OFS FROM OFS WHERE TRIM(ID_OFS) = "+ str(OF)+ "))) ) SELECT LISTAGG(ID_ARTICLE, ', ') || ' : ' || REF_LIBELLE FROM ART_REF GROUP BY REF_LIBELLE")
							article_fils = csr5.fetchall()
							for i in article_fils :
								print("=> ", i[0])
							if DONNEES[1] is None :
								matiere = input("Il n'y a aucune matière renseigné pour l'OF " + OF + " , veuillez la renseigner :  ")
								saisi_matiere = True
							else :
								matiere = DONNEES[1]
							if DONNEES[0] is None :
								epaisseur = input("Il n'y a aucune épaisseur renseigné pour l'OF " + OF + " , veuillez la renseigner :  ")
								saisi_epaisseur = True
							else :
								epaisseur = DONNEES[0]
						if DONNEES[0] is not None :
							epaisseur = DONNEES[0]
						if DONNEES[1] is not None :
							matiere = DONNEES[1]
						if DONNEES[2] is None :
							client = input("Il n'y a aucun Client renseigné pour l'OF " + OF + " , veuillez le renseigner :  ")
							saisi_client = True
						else :
							client = DONNEES[2]
							
						if isinstance(epaisseur, int) :
							epaisseur = epaisseur + ".0"
						epaisseur = epaisseur.replace(",",".")
						print("\n Voici les données : \n OF : "+ OF + "\n Client : " + client + "\n épaisseur : " +epaisseur + " \n matière : "+matiere + "\n numéro de programme : " + dernier_programme + " \n")
						logger.info("\n Voici les données : \n OF : "+ OF + "\n Client : " + client + "\n épaisseur : " +epaisseur + " \n matière : "+matiere + "\n numéro de programme : " + dernier_programme + " \n")
						if client[:8] == "DASSAULT" :
							client = "DASSAULT"
						elif client[:18] == "AIRBUS HELICOPTERS" :
							client = "AIRBUS HELICOPTERS"
						
						colonne = True
						for ligne_DONNEES in sheet_IMP650.iter_rows():
							if len(ligne_DONNEES) >= 6:
								#Récolte des données correspondantes au client matiere et epaisseur
								if matiere == str(ligne_DONNEES[2].value) :
									verif_matiere = True
								else :
									verif_matiere = False
								
								if isinstance(ligne_DONNEES[3].value,int) :
									ligne_DONNEES[3].value = str(ligne_DONNEES[3].value) + ".0"
								if isinstance(ligne_DONNEES[4].value,int) :
									ligne_DONNEES[4].value = str(ligne_DONNEES[4].value) + ".0"
								if isinstance(epaisseur, int) :
									epaisseur = str(epaisseur) + ".0"
								if ligne_DONNEES[3].value == "X" and ligne_DONNEES[4].value == "X" :
									verif_epaisseur = False
									if ligne_DONNEES[0].value == int(dernier_programme) and ligne_DONNEES[1].value == client and str(ligne_DONNEES[2].value) == matiere :
										verif_DONNEES = True
									else : 
										verif_DONNEES = False
								elif ligne_DONNEES[0].value == int(dernier_programme) and ligne_DONNEES[1].value == client and str(ligne_DONNEES[2].value) == matiere and float(ligne_DONNEES[3].value.replace(",",".")) <= float(epaisseur) and float(ligne_DONNEES[4].value.replace(",",".")) >= float(epaisseur):
									verif_epaisseur = True
									verif_DONNEES = True
								else :
									verif_DONNEES = False
									
								#Récupération des données de l'IMP650	
								if verif_DONNEES :
									verif_matiere = True
									epaisseur.replace(".",",")
									contenu_bonne_ligne = ligne_DONNEES
									consigne = contenu_bonne_ligne[5].value
									tolerance = contenu_bonne_ligne[6].value
									temps_montee = contenu_bonne_ligne[7].value
									temps_latence = contenu_bonne_ligne[8].value
									temps_maintien = contenu_bonne_ligne[9].value
									temps_maintien_min = contenu_bonne_ligne[10].value
									temps_maintien_max = contenu_bonne_ligne[11].value
									trempe = contenu_bonne_ligne[12].value
									temperature_min_bac_avant_trempe = contenu_bonne_ligne[13].value
									temperature_max_bac_avant_trempe = contenu_bonne_ligne[14].value
									temperature_max_bac__apres_trempe = contenu_bonne_ligne[15].value
									variation_temperature = contenu_bonne_ligne[16].value
									temps_transfert_imp = contenu_bonne_ligne[17].value
									gradient_descente = contenu_bonne_ligne[18].value
									valeur_gradient = contenu_bonne_ligne[19].value
									temperature_finale = contenu_bonne_ligne[20].value
									groupe_2_open.seek(0)
									groupe_1_open.seek(0)
									
									if verif_epaisseur == True :
										result_epaisseur = True
									
									
									result_matiere = True
									result_client = True
									
									#tfrc_global = Temps point Froid Regulation point Chaud
									tfrc_global = [[],[],[],[]]
									ligne_actuelle = 1
									consigne_min = consigne - tolerance
									consigne_max = consigne + tolerance
									
									#Récupération des données du fichier Group-1
									for row_groupe1 in cr_groupe1 :	
										if row_groupe1[1] == 'Nom=' and row_groupe1[0] == 'Appareil':
											Four = row_groupe1[2]
										if row_groupe1[0] == "Date/Heure" and row_groupe1[1] == '':
											nom_colonne_groupe_1 = row_groupe1
										elif row_groupe1[0] == '' :
											unitee = row_groupe1
										if ligne_actuelle >= 11 :
											if row_groupe1[1] == '' :
												if not "." in row_groupe1[0] :
													tfrc_global[0].append(parse(row_groupe1[0],dayfirst=True))
													tfrc_global[1].append(float(row_groupe1[2].replace(",",".")))
													tfrc_global[2].append(float(row_groupe1[4].replace(",",".")))
													tfrc_global[3].append(float(row_groupe1[6].replace(",",".")))
										ligne_actuelle+=1
									
									
									#on récupère toutes les plus petites données de frc dans le cinquième de la table comme ça ensuite nous pouvons obtenir quand est-ce que le montée commence
									index_debut_montee = round((tfrc_global[1].index(min(tfrc_global[1][:round(len(tfrc_global[1])/5)]))+tfrc_global[2].index(min(tfrc_global[2][:round(len(tfrc_global[2])/5)]))+tfrc_global[3].index(min(tfrc_global[3][:round(len(tfrc_global[3])/5)])))/3)+1
									
									#on cherche le moment ou le premier de frc qui arrive a consigne_min pour commencer la latence
									index_debut_latence = min(tfrc_global[1][index_debut_montee:].index(next(i for i in tfrc_global[1][index_debut_montee:] if i >= consigne_min)), tfrc_global[2][index_debut_montee:].index(next(i for i in tfrc_global[2][index_debut_montee:] if i >= consigne_min)),tfrc_global[3][index_debut_montee:].index(next(i for i in tfrc_global[3][index_debut_montee:] if i >= consigne_min))) + index_debut_montee 
									
									#le debut du maintien commence lorsque frc est au dessus de consigne_min
									index_debut_maintien = max(tfrc_global[1][index_debut_montee:].index(next(i for i in tfrc_global[1][index_debut_montee:] if i >= consigne_min)), tfrc_global[2][index_debut_montee:].index(next(i for i in tfrc_global[2][index_debut_montee:] if i >= consigne_min)),tfrc_global[3][index_debut_montee:].index(next(i for i in tfrc_global[3][index_debut_montee:] if i >= consigne_min))) + index_debut_montee 
									
									#la redescenete commence quand la dernière courbe en partant de la fin ne correspond plus a la consigne_min
									index_redescente = min(len(tfrc_global[1]) - 1 - list(reversed(tfrc_global[1])).index(next(i for count, i in enumerate(list(reversed(tfrc_global[1][index_debut_maintien:])))  if i < consigne_min and list(reversed(tfrc_global[1][index_debut_maintien:]))[count+1] >= consigne_min))+1,len(tfrc_global[2]) - 1 - list(reversed(tfrc_global[2])).index(next(i for count, i in enumerate(list(reversed(tfrc_global[2][index_debut_maintien:]))) if i < consigne_min and list(reversed(tfrc_global[2][index_debut_maintien:]))[count+1] >= consigne_min))+1,len(tfrc_global[3]) - 1 - list(reversed(tfrc_global[3])).index(next(i for count, i in enumerate(list(reversed(tfrc_global[3][index_debut_maintien:]))) if i < consigne_min and list(reversed(tfrc_global[3][index_debut_maintien:]))[count+1] >= consigne_min))-1)
								
									bac_temp = []
									ligne_actuelle = 0
									# vérification des données correspondantes spécialement a la trempe
									if trempe == "OUI":
										result_trempe = True
										# Récupération des données du fichier Group-2
										for row_trempe in cr_groupe2:
											if row_trempe[1] == '' and "." not in row_trempe[0]:
												bac_temp.append(row_trempe[2])
												ligne_actuelle+=1
											elif "Temps de transfert" in row_trempe[1] :
												ligne_transfert = ligne_actuelle 
												# récupéré le temps de transfert
												temps_tranfert = row_trempe[1].split("Temps de transfert ")
												result_temps_tranfert = temps_tranfert[1]
										
										#Récupération des températures du bac
										ligne_actuelle = 0
										Values = []
										#On part de 2 avant car lorsque l'on met la pièce dans bac l'eau vas forcément chauffé imédiatement donc il est préférable de prendre 20 secondes avant que les pieces entre dans le bac pour obtenir la temparture minimum réelle
										result_bac_temp_min = bac_temp[ligne_transfert-2]
										result_bac_temp_min = float(result_bac_temp_min.replace(",","."))
										for k in bac_temp[ligne_transfert-2:] :
											ligne_actuelle+=1
											
											#On prende la valeur la plus petite apres 150 secondes car a ce moment la nous sommes sur que la température du bac max est réellé
											if ligne_actuelle >=15 :
												for ii in bac_temp[ligne_transfert+12:] :
													Values.append(ii)
												result_bac_temp_max = max(Values)
												result_bac_temp_max = float(result_bac_temp_max.replace(",","."))
											
										#Si la trempe n'a pas durée au minimum 150 secondes alors nous prennons la dernière pour se rapprocher le plus possible de la température réelle max
										else : 
											result_bac_temp_max = bac_temp[-1]
											result_bac_temp_max = float(result_bac_temp_max.replace(",","."))
										result_second_after_bac = (len(bac_temp) - ligne_transfert) * 10
											

										#récupération de l'élévation de la température du bac de trempe
										result_elevation_bac_trempe = result_bac_temp_max - result_bac_temp_min
										result_elevation_bac_trempe = round(result_elevation_bac_trempe,1)	
										
										

									#Si il n'y a pas de trempe il y a un gradient (ou pas) alors le client a informé comment il souhaite comment il aimerais que la température baisse toute les heures		
									if  gradient_descente == "OUI" :
										index_valeur_temperature_max_froid = tfrc_global[1].index(next(i for i in tfrc_global[1][index_redescente:] if i <= temperature_finale+5))-1
										index_valeur_temperature_max_regulation = tfrc_global[2].index(next(i for i in tfrc_global[2][index_redescente:] if i <= temperature_finale+5))-1
										index_valeur_temperature_max_chaud = tfrc_global[3].index(next(i for i in tfrc_global[3][index_redescente:] if i <= temperature_finale+5))-1 
										
											
										tfrc_gradient = []
										result_gradient_froid = []
										result_gradient_regulation = []
										result_gradient_chaud = []
										#On récupère comment la température baisse tout les 30 minutes pour voir comment cela diminue
										for i in range(len(tfrc_global[0])):
											tfrc_gradient.append(i)
										for i in range(len(tfrc_global[1][index_redescente:index_valeur_temperature_max_froid-30])):
											result_gradient_froid.append((tfrc_global[1][index_redescente+i+30]-tfrc_global[1][index_redescente+i]))
										for i in range(len(tfrc_global[2][index_redescente:index_valeur_temperature_max_regulation-30])):
											result_gradient_regulation.append((tfrc_global[2][index_redescente+i+30]-tfrc_global[2][index_redescente+i]))
										for i in range(len(tfrc_global[3][index_redescente:index_valeur_temperature_max_chaud-30])):
											result_gradient_chaud.append((tfrc_global[3][index_redescente+i+30]-tfrc_global[3][index_redescente+i]))
									
									#Le client n'a rien demandé alors on ne fais rien
									if gradient_descente == "X" and trempe == "X":
										print("Aucun gradient informé et aucune trempe informé")
									
									if not(tfrc_global[1][index_debut_maintien] >= consigne_min and tfrc_global[2][index_debut_maintien] and tfrc_global[3][index_debut_maintien]):
										print("ATTENTION DEBUT MAINTIEN FAUSSÉ")
								
									ax.plot(tfrc_global[0],tfrc_global[1],label = nom_colonne_groupe_1[2])		# drawing of the curve for cold point temperatures
									ax.plot(tfrc_global[0],tfrc_global[2],label = nom_colonne_groupe_1[4])		# drawing of the curve for regulation temperatures
									ax.plot(tfrc_global[0],tfrc_global[3],label = nom_colonne_groupe_1[6])		# drawing of the curve for hot point temperatures
		
									plt.title(f"{fournee} n° {OF} edité le: {jour}/{mois}/{annee}")
									
									
									
									
									# Drawing "the red curve" to help debuging

									ax.set_ylabel('Température ' + unitee[2])
									legend = ax.legend(loc='lower left', shadow=True, fontsize='x-large')

									# Draw vertical lines to markup events:
									ax.axvline(x=tfrc_global[0][index_debut_montee],color = 'blue')						# - Blue is for loading
									ax.axvline(x=tfrc_global[0][index_debut_latence],color = 'limegreen')				# - LimeGeen is for treatment begening
									ax.axvline(x=tfrc_global[0][index_redescente],color = 'red')						# - Red is for treatment end
									ax.axvline(x=tfrc_global[0][index_debut_maintien], color = 'purple')
									ax.axhline(y=consigne_min, xmin=0.0, xmax=1.0, color='black')

					
									plt.tight_layout()	# Put a nicer background color on the legend.
									
									
									
									
									
									#Récupération du point froid minimale et maximale, de la régulation minimale et maximale et du point chaud minimale et maximale du groupe-1
									result_pt_froid_min = min(tfrc_global[1][index_debut_maintien:index_redescente-1])
									result_pt_froid_max = max(tfrc_global[1][index_debut_maintien:index_redescente-1])
									
									result_regulation_min = min(tfrc_global[2][index_debut_maintien:index_redescente-1])
									result_regulation_max = max(tfrc_global[2][index_debut_maintien:index_redescente-1])	
									
									result_pt_chaud_min = min(tfrc_global[3][index_debut_maintien:index_redescente-1])
									result_pt_chaud_max = max(tfrc_global[3][index_debut_maintien:index_redescente-1])
									
									
									
									#création des booléen pour :
									if "X" == temps_montee :
										temps_montee = "-"
									if "X" == temps_latence :
										temps_latence = "-"
									if "X" == temps_maintien_min :
										temps_maintien = "-"
									if "X" == temps_maintien_max :
										temps_maintien_max = "-"
									
									#Récupération des temps pour le maintien, la montée, la latence
									result_temps_maintien = ((index_redescente-1) - index_debut_maintien) * 10 	
									result_temps_maintien = strftime('%H:%M:%S', gmtime(result_temps_maintien))
									
									result_temps_montee = (index_debut_latence - index_debut_montee) * 10
									result_temps_montee = strftime('%H:%M:%S', gmtime(result_temps_montee))
									
									
									result_temps_latence = (index_debut_maintien - index_debut_latence ) * 10
									result_temps_latence = strftime('%H:%M:%S', gmtime(result_temps_latence))
									
									
									if temps_maintien_max == "-" :
										hms_temps_maintien_max = "-"
									else : 
										hms_temps_maintien_max = strftime('%H:%M:%S', gmtime(temps_maintien_max*60))
									
									if temps_maintien_min == "-" :
										hms_temps_maintien_min = "-"
									else : 
										hms_temps_maintien_min = strftime('%H:%M:%S', gmtime(temps_maintien_min*60))
										
									if temps_montee == "-" :
										hms_temps_montee = "-"
									else : 
										hms_temps_montee = strftime('%H:%M:%S', gmtime(temps_montee*60))
									
									if temps_latence == "-" :
										hms_temps_latence = "-"
									else : 
										hms_temps_latence = strftime('%H:%M:%S', gmtime(temps_latence*60))
									
									# Vérification des données pour les conformités
									if result_regulation_max > consigne_max :
									
										conform_regulation_max = "NON OK"
									else :
										conform_regulation_max = "OK"
									if result_regulation_min < consigne_min :
									
										conform_regulation_min = "NON OK"
									else :
										conform_regulation_min = "OK"
										
									if result_pt_froid_max > consigne_max :
										conform_pt_froid_max = "NON OK"
									else :
										conform_pt_froid_max = "OK"
									if result_pt_froid_min < consigne_min :
									
										conform_pt_froid_min = "NON OK"
									else :
										conform_pt_froid_min = "OK"
									if result_pt_chaud_max > consigne_max :
										conform_pt_chaud_max = "NON OK"
									else :
										conform_pt_chaud_max = "OK"
									if result_pt_chaud_min < consigne_min :
									
										conform_pt_chaud_min = "NON OK"
									else :
										conform_pt_chaud_min = "OK"
									
									#utilisation de la fonction verification_temps pour savoir si le temps est entre les exigences du client	
									conform_temps_maintien = comparaison_temps(result_temps_maintien,temps_maintien_max,temps_maintien_min)
									conform_temps_latence = comparaison_temps(result_temps_latence,temps_latence,"-")
									conform_temps_montee = comparaison_temps(result_temps_montee, temps_montee,"-")
								
									#Transformation en fonction de la validité de la donnée en NON OK ou OK ou "-" si une trempe est faite
									if trempe == "OUI" :
										if temperature_min_bac_avant_trempe == "X" :
											temperature_min_bac_avant_trempe = "-"
										elif result_bac_temp_min >= temperature_min_bac_avant_trempe and result_bac_temp_min <= temperature_max_bac_avant_trempe :
											conform_temp_bac_avant_trempe = "OK"
										else :
											conform_temp_bac_avant_trempe = "NON OK"
										
										if result_bac_temp_max<= temperature_max_bac__apres_trempe :
											conform_temp_bac_apres_trempe = "OK"
										else :
											conform_temp_bac_apres_trempe = "NON OK"
										if result_elevation_bac_trempe <= variation_temperature :
											conform_temp_bac_elevation = "OK"
										else :
											conform_temp_bac_elevation = "NON OK"
										if result_temps_tranfert.split(":")[2] <= temps_tranfert[1].split(":")[2] :
											conform_temps_transfert = "OK"
										else :
											conform_temps_transfert = "NON OK"
											
									#Transformation en fonction de la validité de la donnée en NON OK ou OK ou "-" si un gradient est fait
									if gradient_descente == "OUI" :
										if round(min(result_gradient_froid)*12,1) >= valeur_gradient :
											conform_gradient_pt_froid = "OK"
										else : 
											conform_gradient_pt_froid = "NON OK"
										if round(min(result_gradient_regulation)*12,1) >= valeur_gradient :
											conform_gradient_pt_regulation = "OK"
										else : 
											conform_gradient_pt_regulation = "NON OK"
										if round(min(result_gradient_chaud)*12,1) >= valeur_gradient :
											conform_gradient_pt_chaud = "OK"
										else : 
											conform_gradient_pt_chaud = "NON OK"
								
									
									#Écriture du noms des lignes																				Écriture des données récupéré dans le four																							Écriture des éxigence du client																																		Écriture de la validité des données récupérées en fonction des éxigence
									x = 0.55 + 0.30
									y = 0.80
									plt.text(x:=x-0.30,y:=y-0.04,"Programme n°" + str(dernier_programme), transform=ax.transAxes); 				plt.text(x:=x+0.15,y,"Mesuré", transform=ax.transAxes); 																			plt.text(x:=x+0.075,y,"Exigé", transform=ax.transAxes); 																											plt.text(x:=x+0.075,y,"Conformité", transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.0170,"Tolérance +/-" + str(tolerance), transform=ax.transAxes); 					plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					plt.text(x:=x+0.075,y,"", transform=ax.transAxes); 																													plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"Régulation Max :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y, str(result_regulation_max) + "°C", transform=ax.transAxes); 													plt.text(x:=x+0.075,y, str(consigne_max) + "°C", transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_regulation_max,color = verif_color(conform_regulation_max), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"Régulation Min :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y, str(result_regulation_min) + "°C", transform=ax.transAxes); 													plt.text(x:=x+0.075,y, str(consigne_min) + "°C", transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_regulation_min,color = verif_color(conform_regulation_min), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"PT froid Max :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y, str(result_pt_froid_max) + "°C", transform=ax.transAxes); 													plt.text(x:=x+0.075,y, str(consigne_max) + "°C", transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_pt_froid_max,color = verif_color(conform_pt_froid_max), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"PT froid Min :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y, str(result_pt_froid_min) + "°C", transform=ax.transAxes); 													plt.text(x:=x+0.075,y, str(consigne_min) + "°C", transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_pt_froid_min, color = verif_color(conform_pt_froid_min), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"PT chaud Max :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y, str(result_pt_chaud_max) + "°C", transform=ax.transAxes); 													plt.text(x:=x+0.075,y, str(consigne_max) + "°C", transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_pt_chaud_max,color = verif_color(conform_pt_chaud_max), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"PT chaud Min :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y, str(result_pt_chaud_min) + "°C", transform=ax.transAxes); 													plt.text(x:=x+0.075,y, str(consigne_min) + "°C", transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_pt_chaud_min,color = verif_color(conform_pt_chaud_min), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"", transform=ax.transAxes); 													plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					plt.text(x:=x+0.075,y, "Min :", transform=ax.transAxes); 																											plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.02,"Temps de maintien :", transform=ax.transAxes); 								plt.text(x:=x+0.15,y,result_temps_maintien, transform=ax.transAxes); 																plt.text(x:=x+0.075,y, hms_temps_maintien_min, transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,conform_temps_maintien,color = verif_color(conform_temps_maintien), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.0170,"", transform=ax.transAxes); 												plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					plt.text(x:=x+0.075,y, "Max :", transform=ax.transAxes); 																											plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.0170,"", transform=ax.transAxes); 												plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					plt.text(x:=x+0.075,y, hms_temps_maintien_max, transform=ax.transAxes); 																							plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"Temps de latence :", transform=ax.transAxes); 								plt.text(x:=x+0.15,y,result_temps_latence, transform=ax.transAxes); 																plt.text(x:=x+0.075,y, hms_temps_latence, transform=ax.transAxes); 																									plt.text(x:=x+0.075,y,conform_temps_latence,color = verif_color(conform_temps_latence), transform=ax.transAxes);
									plt.text(x:=x-0.30,y:=y-0.04,"Temps de montée :", transform=ax.transAxes); 									plt.text(x:=x+0.15,y,result_temps_montee, transform=ax.transAxes); 																	plt.text(x:=x+0.075,y, hms_temps_montee, transform=ax.transAxes); 																									plt.text(x:=x+0.075,y,conform_temps_montee,color = verif_color(conform_temps_montee), transform=ax.transAxes);
									if trempe == "OUI" :
										plt.text(x:=x-0.30,y:=y-0.04,"Temp bac init :", transform=ax.transAxes); 								plt.text(x:=x+0.15,y,str(result_bac_temp_min) + "°C", transform=ax.transAxes); 														plt.text(x:=x+0.075,y,"[" +str(temperature_min_bac_avant_trempe) +"°C;" +str(temperature_max_bac_avant_trempe) +"°C]", transform=ax.transAxes); 					plt.text(x:=x+0.075,y,conform_temp_bac_avant_trempe,color = verif_color(conform_temp_bac_avant_trempe), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"Temp bac fin :", transform=ax.transAxes); 								plt.text(x:=x+0.15,y,str(result_bac_temp_max) + "°C", transform=ax.transAxes); 														plt.text(x:=x+0.075,y,str(temperature_max_bac__apres_trempe) + "°C" , transform=ax.transAxes); 																		plt.text(x:=x+0.075,y,conform_temp_bac_apres_trempe,color = verif_color(conform_temp_bac_apres_trempe), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"Élévation temp bac :", transform=ax.transAxes); 							plt.text(x:=x+0.15,y,result_elevation_bac_trempe, transform=ax.transAxes); 															plt.text(x:=x+0.075,y, str(variation_temperature) + "°C" , transform=ax.transAxes); 																				plt.text(x:=x+0.075,y,conform_temp_bac_elevation,color = verif_color(conform_temp_bac_elevation), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.0170,"Relevé après :", transform=ax.transAxes); 								plt.text(x:=x+0.15,y,str(result_second_after_bac) + "''", transform=ax.transAxes); 													plt.text(x:=x+0.075,y,"", transform=ax.transAxes); 																													plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"Temps de transfert :", transform=ax.transAxes); 							plt.text(x:=x+0.15,y,result_temps_tranfert, transform=ax.transAxes); 																plt.text(x:=x+0.075,y, str(temps_transfert_imp) + "''", transform=ax.transAxes); 																					plt.text(x:=x+0.075,y,conform_temps_transfert,color = verif_color(conform_temps_transfert), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"", transform=ax.transAxes); 												plt.text(x:=x+0.15,y,"(Pas de gradient effectué)",color = "blue", transform=ax.transAxes); 											plt.text(x:=x+0.075,y,"", transform=ax.transAxes); 																													plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									
									elif gradient_descente == "OUI" :
										plt.text(x:=x-0.30,y:=y-0.04,"PT froid gradient :", transform=ax.transAxes); 							plt.text(x:=x+0.15,y,str(round(min(result_gradient_froid)*12,1)) + "°C", transform=ax.transAxes); 									plt.text(x:=x+0.075,y, str(valeur_gradient) + "°C", transform=ax.transAxes); 																						plt.text(x:=x+0.075,y,conform_gradient_pt_froid,color =verif_color(conform_gradient_pt_froid), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"PT regulation gradient :", transform=ax.transAxes); 						plt.text(x:=x+0.15,y,str(round(min(result_gradient_regulation)*12,1))+ "°C", transform=ax.transAxes); 								plt.text(x:=x+0.075,y, str(valeur_gradient) + "°C", transform=ax.transAxes); 																						plt.text(x:=x+0.075,y,conform_gradient_pt_regulation,color =verif_color(conform_gradient_pt_regulation), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"PT chaud gradient :", transform=ax.transAxes); 							plt.text(x:=x+0.15,y,str(round(min(result_gradient_chaud)*12,1))+ "°C", transform=ax.transAxes); 									plt.text(x:=x+0.075,y, str(valeur_gradient) + "°C", transform=ax.transAxes); 																						plt.text(x:=x+0.075,y,conform_gradient_pt_chaud,color =verif_color(conform_gradient_pt_chaud), transform=ax.transAxes);
										plt.text(x:=x-0.30,y:=y-0.04,"", transform=ax.transAxes); 												plt.text(x:=x+0.15,y,"(PAS de trempe effectué)",color= "blue", transform=ax.transAxes); 											plt.text(x:=x+0.075,y, "", transform=ax.transAxes); 																												plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									else :
										plt.text(x:=x-0.30,y:=y-0.04,"", transform=ax.transAxes); 												plt.text(x:=x+0.15,y,"(Pas de trempe effectué) (Pas de gradient effectué)",color = "blue", transform=ax.transAxes); 					plt.text(x:=x+0.075,y, "", transform=ax.transAxes); 																											plt.text(x:=x+0.075,y,"", transform=ax.transAxes);
									if changement_programme :
										plt.text(0.2, 0.1,"ATTENTION IL Y A EU UN CHANGEMENT DE PROGRAMME", transform = ax.transAxes);
									
									
									#Regarde sur la validation des données CLIENT MATIERE EPAISSEUR	
									if client != ligne_DONNEES[1].value :
										result_client = False
									if matiere != str(ligne_DONNEES[2].value) :
										result_matiere = False
									if verif_epaisseur == True and 'Épaisseur Min (en mm)' not in ligne_DONNEES[3].value  :
										if float(ligne_DONNEES[3].value.replace(",",".")) > float(epaisseur) and float(ligne_DONNEES[4].value.replace(",",".")) < float(epaisseur) :
											result_epaisseur = False
									result_donnees = Four + "\nMatière sur OF :\n =>" + matiere + " - Epaisseur : " + epaisseur + " (MM)"
									
									plt.text(0.18,0.07,result_donnees,transform=ax.transAxes)
									if saisi_matiere :
										if result_matiere :
											plt.text(0.2,0.055,"Matière OK",color="blue",transform=ax.transAxes)
										else :
											plt.text(0.2,0.055,"Matière NON OK",color="blue",transform=ax.transAxes)
									else :
										if result_matiere :
											plt.text(0.2,0.055,"Matière OK",color="green",transform=ax.transAxes)
										else :
											plt.text(0.2,0.055,"Matière NON OK",color="red",transform=ax.transAxes)
											
									if saisi_epaisseur :
										if verif_epaisseur :
											if result_epaisseur :
												plt.text(0.2,0.04,"Épaisseur OK",color="blue",transform=ax.transAxes)
											else :
												plt.text(0.2,0.04,"Épaisseur NON OK",color="blue",transform=ax.transAxes)	
										else :
											plt.text(0.2,0.04,"",transform=ax.transAxes)
									else :
										if verif_epaisseur :
											if result_epaisseur :
												plt.text(0.2,0.04,"Épaisseur OK",color="green",transform=ax.transAxes)
											else :
												plt.text(0.2,0.04,"Épaisseur NON OK",color="red",transform=ax.transAxes)	
										else :
											plt.text(0.2,0.04,"",transform=ax.transAxes)
											
									if saisi_client :
										if result_client :
												plt.text(0.2,0.0225,"CLIENT : Client OK",color="blue",transform=ax.transAxes)
												
										else :
											plt.text(0.2,0.0225,"Client : Client NON OK",color="blue",transform=ax.transAxes)
									else :
										if result_client :
											plt.text(0.2,0.0225,"CLIENT : Client OK",color="green",transform=ax.transAxes)
											
										else :
											plt.text(0.2,0.0225,"Client : Client NON OK",color="red",transform=ax.transAxes)
									plt.text(0.18,0.0015,"=>" + str(client),transform=ax.transAxes)
									break
									
						if verif_DONNEES == False:
							groupe_2_open.seek(0)
							groupe_1_open.seek(0)
							tfrc_global = [[],[],[],[]]
							ligne_actuelle = 1
							
							if mode_auto == 1 :
								tolerance = int(input("Veuillez saisir la tolérance pour l'OF "+ OF+" :  "))
							else :
								tolerance = 5
							
							consigne_min = dernier_consigne - tolerance
							consigne_max = dernier_consigne + tolerance
							
							#Récupération des données du fichier Group-1
							for row_groupe1 in cr_groupe1 :	
								if row_groupe1[0] == "Date/Heure" and row_groupe1[1] == '':
									nom_colonne_groupe_1 = row_groupe1
								elif row_groupe1[0] == '' and row_groupe1[1] == '' :
									unitee = row_groupe1
								if ligne_actuelle >= 11 :
									if row_groupe1[1] == '' :
										if not "." in row_groupe1[0] :
											tfrc_global[0].append(parse(row_groupe1[0],dayfirst=True))
											tfrc_global[1].append(float(row_groupe1[2].replace(",",".")))
											tfrc_global[2].append(float(row_groupe1[4].replace(",",".")))
											tfrc_global[3].append(float(row_groupe1[6].replace(",",".")))
								ligne_actuelle+=1
								
							ax.plot(tfrc_global[0],tfrc_global[1],label = nom_colonne_groupe_1[2])		# drawing of the curve for cold point temperatures
							ax.plot(tfrc_global[0],tfrc_global[2],label = nom_colonne_groupe_1[4])		# drawing of the curve for regulation temperatures
							ax.plot(tfrc_global[0],tfrc_global[3],label = nom_colonne_groupe_1[6])		# drawing of the curve for hot point temperatures

							plt.title(f"{fournee} n° {OF} edité le: {jour}/{mois}/{annee}")
							
							#on récupère toutes les plus petites données de frc dans le cinquième de la table comme ça ensuite nous pouvons obtenir quand est-ce que le montée commence
							index_debut_montee = round((tfrc_global[1].index(min(tfrc_global[1][:round(len(tfrc_global[1])/5)]))+tfrc_global[2].index(min(tfrc_global[2][:round(len(tfrc_global[2])/5)]))+tfrc_global[3].index(min(tfrc_global[3][:round(len(tfrc_global[3])/5)])))/3)+1
							
							#on cherche le moment ou le premier de frc qui arrive a consigne_min pour commencer la latence
							index_debut_latence = min(tfrc_global[1][index_debut_montee:].index(next(i for i in tfrc_global[1][index_debut_montee:] if i >= consigne_min)), tfrc_global[2][index_debut_montee:].index(next(i for i in tfrc_global[2][index_debut_montee:] if i >= consigne_min)),tfrc_global[3][index_debut_montee:].index(next(i for i in tfrc_global[3][index_debut_montee:] if i >= consigne_min))) + index_debut_montee 
							
							#le debut du maintien commence lorsque frc est au dessus de consigne_min
							index_debut_maintien = max(tfrc_global[1][index_debut_montee:].index(next(i for i in tfrc_global[1][index_debut_montee:] if i >= consigne_min)), tfrc_global[2][index_debut_montee:].index(next(i for i in tfrc_global[2][index_debut_montee:] if i >= consigne_min)),tfrc_global[3][index_debut_montee:].index(next(i for i in tfrc_global[3][index_debut_montee:] if i >= consigne_min))) + index_debut_montee 
							
							#la redescenete commence quand la dernière courbe en partant de la fin ne correspond plus a la consigne_min
							index_redescente = min(len(tfrc_global[1]) - 1 - list(reversed(tfrc_global[1])).index(next(i for count, i in enumerate(list(reversed(tfrc_global[1][index_debut_maintien:])))  if i < consigne_min and list(reversed(tfrc_global[1][index_debut_maintien:]))[count+1] >= consigne_min))+1,len(tfrc_global[2]) - 1 - list(reversed(tfrc_global[2])).index(next(i for count, i in enumerate(list(reversed(tfrc_global[2][index_debut_maintien:]))) if i < consigne_min and list(reversed(tfrc_global[2][index_debut_maintien:]))[count+1] >= consigne_min))+1,len(tfrc_global[3]) - 1 - list(reversed(tfrc_global[3])).index(next(i for count, i in enumerate(list(reversed(tfrc_global[3][index_debut_maintien:]))) if i < consigne_min and list(reversed(tfrc_global[3][index_debut_maintien:]))[count+1] >= consigne_min))-1)

							#Récupération du point froid minimale et maximale, de la régulation minimale et maximale et du point chaud minimale et maximale du groupe-1
							result_pt_froid_min = min(tfrc_global[1][index_debut_maintien:index_redescente-1])
							result_pt_froid_max = max(tfrc_global[1][index_debut_maintien:index_redescente-1])
							
							result_regulation_min = min(tfrc_global[2][index_debut_maintien:index_redescente-1])
							result_regulation_max = max(tfrc_global[2][index_debut_maintien:index_redescente-1])	
							
							result_pt_chaud_min = min(tfrc_global[3][index_debut_maintien:index_redescente-1])
							result_pt_chaud_max = max(tfrc_global[3][index_debut_maintien:index_redescente-1])
							
							#Récupération des temps pour le maintien, la montée, la latence
							result_temps_maintien = ((index_redescente-1) - index_debut_maintien) * 10 	
							result_temps_maintien = strftime('%H:%M:%S', gmtime(result_temps_maintien))
							
							result_temps_montee = (index_debut_latence - index_debut_montee) * 10
							result_temps_montee = strftime('%H:%M:%S', gmtime(result_temps_montee))
							
							
							result_temps_latence = (index_debut_maintien - index_debut_latence ) * 10
							result_temps_latence = strftime('%H:%M:%S', gmtime(result_temps_latence))
							
							# Drawing "the red curve" to help debuging
							ax.set_ylabel('Température ' + unitee[2])
							legend = ax.legend(loc='lower left', shadow=True, fontsize='x-large')
							
							# Draw vertical lines to markup events:
							ax.axvline(x=tfrc_global[0][index_debut_montee],color = 'blue')						# - Blue is for loading
							ax.axvline(x=tfrc_global[0][index_debut_latence],color = 'limegreen')				# - LimeGeen is for treatment begening
							ax.axvline(x=tfrc_global[0][index_redescente],color = 'red')						# - Red is for treatment end
							ax.axvline(x=tfrc_global[0][index_debut_maintien], color = 'purple')
							ax.axhline(y=consigne_min, xmin=0.0, xmax=1.0, color='black')
							plt.tight_layout()	# Put a nicer background color on the legend.
							
							
							bac_temp = []
							ligne_actuelle = 0
							result_trempe = False
							verif_gradient = False
							# Récupération des données du fichier Group-2
							for row_trempe in cr_groupe2:
								if row_trempe[1] == '' and "." not in row_trempe[0]:
									bac_temp.append(row_trempe[2])
									ligne_actuelle+=1
								elif "Temps de transfert" in row_trempe[1] :
									result_trempe = True
									ligne_transfert = ligne_actuelle 
									# récupéré le temps de transfert
									temps_tranfert = row_trempe[1].split("Temps de transfert ")
									result_temps_tranfert = temps_tranfert[1]
							# pdb.set_trace()	
							# vérification des données correspondantes spécialement a la trempe
							if result_trempe:
								#Récupération des températures du bac
								ligne_actuelle = 0
								Values = []
								#On part de 2 avant car lorsque l'on met la pièce dans bac l'eau vas forcément chauffé imédiatement donc il est préférable de prendre 20 secondes avant que les pieces entre dans le bac pour obtenir la temparture minimum réelle
								result_bac_temp_min = bac_temp[ligne_transfert-2]
								result_bac_temp_min = float(result_bac_temp_min.replace(",","."))
								for k in bac_temp[ligne_transfert-2:] :
									ligne_actuelle+=1
									
									#On prende la valeur la plus petite apres 150 secondes car a ce moment la nous sommes sur que la température du bac max est réellé
									if ligne_actuelle >=15 :
										for ii in bac_temp[ligne_transfert+12:] :
											Values.append(ii)
										result_bac_temp_max = max(Values)
										result_bac_temp_max = float(result_bac_temp_max.replace(",","."))
									
								#Si la trempe n'a pas durée au minimum 150 secondes alors nous prennons la dernière pour se rapprocher le plus possible de la température réelle max
								else : 
									result_bac_temp_max = bac_temp[-1]
									result_bac_temp_max = float(result_bac_temp_max.replace(",","."))
								result_second_after_bac = (len(bac_temp) - ligne_transfert) * 10
									

								#récupération de l'élévation de la température du bac de trempe
								result_elevation_bac_trempe = result_bac_temp_max - result_bac_temp_min
								result_elevation_bac_trempe = round(result_elevation_bac_trempe,1)	
							elif min(tfrc_global[1][index_debut_maintien:index_redescente])>255 and min(tfrc_global[1][index_redescente:])<=255 :
								verif_gradient = True
								if min(tfrc_global[1][index_redescente:]) <= 255 :
									index_valeur_temperature_max_froid = tfrc_global[1].index(next(i for i in tfrc_global[1][index_redescente:] if i <= 250+5))-1
								else :
									index_valeur_temperature_max_froid = tfrc_global[1][-1]
								if min(tfrc_global[2][index_redescente:]) <= 255 :
									index_valeur_temperature_max_regulation = tfrc_global[2].index(next(i for i in tfrc_global[2][index_redescente:] if i <= 250+5))-1
								else :
									index_valeur_temperature_max_regulation = tfrc_global[2][-1]
								if min(tfrc_global[3][index_redescente:]) <= 255 :
									index_valeur_temperature_max_chaud = tfrc_global[3].index(next(i for i in tfrc_global[3][index_redescente:] if i <= 250+5))-1 
								else :
									index_valeur_temperature_max_chaud = tfrc_global[3][-1]
								
								tfrc_gradient = []
								result_gradient_froid = []
								result_gradient_regulation = []
								result_gradient_chaud = []
								
								#On récupère comment la température baisse tout les 30 minutes pour voir comment cela diminue
								for i in range(len(tfrc_global[0])):
									tfrc_gradient.append(i)
								for i in range(len(tfrc_global[1][index_redescente:index_valeur_temperature_max_froid-30])):
									result_gradient_froid.append((tfrc_global[1][index_redescente+i+30]-tfrc_global[1][index_redescente+i]))
								for i in range(len(tfrc_global[2][index_redescente:index_valeur_temperature_max_regulation-30])):
									result_gradient_regulation.append((tfrc_global[2][index_redescente+i+30]-tfrc_global[2][index_redescente+i]))
								for i in range(len(tfrc_global[3][index_redescente:index_valeur_temperature_max_chaud-30])):
									result_gradient_chaud.append((tfrc_global[3][index_redescente+i+30]-tfrc_global[3][index_redescente+i]))
									
							#Écriture du noms des lignes																														Écriture des données récupéré dans le four																							
							x = 0.55 + 0.30
							y = 0.80
							plt.text(x:=x-0.30,y:=y-0.04,"Programme n°" + str(dernier_programme), transform=ax.transAxes); 														plt.text(x:=x+0.15,y,"Mesuré", transform=ax.transAxes); 																			
							plt.text(x:=x-0.15,y:=y-0.0170,"Tolérance +/- INFORMÉ MANUELLEMENT OU PAR DEFAUT (5) :" + str(tolerance), transform=ax.transAxes); 					plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					
							plt.text(x:=x-0.15,y:=y-0.04,"Régulation Max :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y, str(result_regulation_max) + "°C", transform=ax.transAxes); 													
							plt.text(x:=x-0.15,y:=y-0.04,"Régulation Min :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y, str(result_regulation_min) + "°C", transform=ax.transAxes); 													
							plt.text(x:=x-0.15,y:=y-0.04,"PT froid Max :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y, str(result_pt_froid_max) + "°C", transform=ax.transAxes); 													
							plt.text(x:=x-0.15,y:=y-0.04,"PT froid Min :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y, str(result_pt_froid_min) + "°C", transform=ax.transAxes); 													
							plt.text(x:=x-0.15,y:=y-0.04,"PT chaud Max :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y, str(result_pt_chaud_max) + "°C", transform=ax.transAxes); 													
							plt.text(x:=x-0.15,y:=y-0.04,"PT chaud Min :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y, str(result_pt_chaud_min) + "°C", transform=ax.transAxes); 													
							plt.text(x:=x-0.15,y:=y-0.04,"", transform=ax.transAxes); 																							plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					
							plt.text(x:=x-0.15,y:=y-0.02,"Temps de maintien :", transform=ax.transAxes); 																		plt.text(x:=x+0.15,y,result_temps_maintien, transform=ax.transAxes); 																
							plt.text(x:=x-0.15,y:=y-0.0170,"", transform=ax.transAxes); 																						plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					
							plt.text(x:=x-0.15,y:=y-0.0170,"", transform=ax.transAxes); 																						plt.text(x:=x+0.15,y,"", transform=ax.transAxes); 																					
							plt.text(x:=x-0.15,y:=y-0.04,"Temps de latence :", transform=ax.transAxes); 																		plt.text(x:=x+0.15,y,result_temps_latence, transform=ax.transAxes); 																
							plt.text(x:=x-0.15,y:=y-0.04,"Temps de montée :", transform=ax.transAxes); 																			plt.text(x:=x+0.15,y,result_temps_montee, transform=ax.transAxes); 																	
							if result_trempe :
								plt.text(x:=x-0.15,y:=y-0.04,"Temp bac init :", transform=ax.transAxes); 																		plt.text(x:=x+0.15,y,str(result_bac_temp_min) + "°C", transform=ax.transAxes); 														
								plt.text(x:=x-0.15,y:=y-0.04,"Temp bac fin :", transform=ax.transAxes); 																		plt.text(x:=x+0.15,y,str(result_bac_temp_max) + "°C", transform=ax.transAxes); 														
								plt.text(x:=x-0.15,y:=y-0.04,"Élévation temp bac :", transform=ax.transAxes); 																	plt.text(x:=x+0.15,y,result_elevation_bac_trempe, transform=ax.transAxes); 															
								plt.text(x:=x-0.15,y:=y-0.0170,"Relevé après :", transform=ax.transAxes); 																		plt.text(x:=x+0.15,y,str(result_second_after_bac) + "''", transform=ax.transAxes); 													
								plt.text(x:=x-0.15,y:=y-0.04,"Temps de transfert :", transform=ax.transAxes); 																	plt.text(x:=x+0.15,y,result_temps_tranfert, transform=ax.transAxes); 																
								plt.text(x:=x-0.15,y:=y-0.04,"", transform=ax.transAxes); 																						plt.text(x:=x+0.15,y,"(Pas de gradient effectué)",color = "blue", transform=ax.transAxes); 											
							elif verif_gradient == True :
								plt.text(x:=x-0.15,y:=y-0.04,"PT froid gradient :", transform=ax.transAxes); 																	plt.text(x:=x+0.15,y,str(round(min(result_gradient_froid)*12,1)) + "°C", transform=ax.transAxes); 									
								plt.text(x:=x-0.15,y:=y-0.04,"PT regulation gradient :", transform=ax.transAxes); 																plt.text(x:=x+0.15,y,str(round(min(result_gradient_regulation)*12,1))+ "°C", transform=ax.transAxes); 								
								plt.text(x:=x-0.15,y:=y-0.04,"PT chaud gradient :", transform=ax.transAxes); 																	plt.text(x:=x+0.15,y,str(round(min(result_gradient_chaud)*12,1))+ "°C", transform=ax.transAxes); 									
								plt.text(x:=x-0.15,y:=y-0.04,"", transform=ax.transAxes); 								plt.text(x:=x+0.15,y,"(PAS de trempe effectué)",color= "blue", transform=ax.transAxes); 											
							if changement_programme :
								plt.text(0.2, 0.08,"ATTENTION IL Y A EU UN CHANGEMENT DE PROGRAMME", transform = ax.transAxes);
							plt.text(0.2,0.055,"CLIENT : " + str(client),transform=ax.transAxes)
							plt.text(0.2,0.04,"Epaisseur : " + str(epaisseur),transform=ax.transAxes)
							plt.text(0.2,0.0225,"Matiere : " + str(matiere),transform=ax.transAxes)
							
						#Fermeture et sauvearde du fichier PDF
						pdf.savefig()  
						plt.close()	
					
					csr3.execute("SELECT N_LIGNE FROM OFS_LISTE_FIC WHERE CD_OFS = "+ str(cd_ofs[0]) +"ORDER BY N_LIGNE DESC") # Regarde dans la table des fichier lie a l'OF combien de document sont lies
					row3 = csr3.fetchone()
					
					if row3:
						ligne = row3[0] + 1
					else : ligne = 1
					
					with open(nom_pdf, 'rb') as f: blob = f.read()
					insert_stmt = (
						"INSERT INTO OFS_LISTE_FIC (CD_OFS, N_LIGNE, NOM_FICHIER,FICHIER,INACTIF) " + "VALUES (?,?,?,?,?)" # cree dans la table des fichiers lies a l'OF un enregistrement issus du four (a la suite des documents deja lies) 
					)
					stringtopush = "'" + nom_pdf + "'"

					data = (str(cd_ofs[0]), str(ligne), stringtopush, blob, '0')
					
					logger.info('pdf edited for OF ' + OF)
					# if not TEST_MODE:
						# csr4.execute(insert_stmt,data)
						# cnx.commit()
						# logger.info('pdf pushed on Oracle for OF ' + OF)
						# os.remove(nom_pdf)

					print("done")
				
				groupe_1_open.close()
				groupe_2_open.close()
				wb.close()	
				
				#Déplacement des fichiers groupe-1 et groupe-2
				if move_after_verif:
					source_path = os.path.join(current_directory, groupe_1)
					destination_path = os.path.join(destination_directory, groupe_1)
					print("Moving from", source_path, "to", destination_path)
					shutil.move(source_path, destination_path)

					source_path = os.path.join(current_directory, groupe_2)
					destination_path = os.path.join(destination_directory, groupe_2)
					print("Moving from", source_path, "to", destination_path)
					shutil.move(source_path, destination_path)
					
		
		else :
			if fin_programme <1 :
				print("fin de programme, plus de fichier a analysé")
				logger.info("Fin de programme, plus de fichier a analysé")
			fin_programme +=1

	except FileNotFoundError as err:
		line_number = traceback.extract_tb(sys.exc_info()[2])[-1][1]  # Récupère le numéro de ligne de l'exception
		logger.error(f"Erreur à la ligne {line_number}: {str(err)}")
		continue

	except Exception as exc:
		line_number = traceback.extract_tb(sys.exc_info()[2])[-1][1]  # Récupère le numéro de ligne de l'exception
		logger.error(f"Erreur à la ligne {line_number}: {str(exc)}")
		continue
os.remove("IMP650 (Critères de validation de traitementthermique).xlsx")
print("Fin de programme dans 10 secondes")
sleep(10)