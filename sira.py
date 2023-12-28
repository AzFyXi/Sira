import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<17:
		speak("Bonjour")
	else:
		speak("Bonsoir")

	assname =("Sira")
	speak("Je suis votre assistante personnel")
	speak(assname)
	

def username():
	speak("Comment dois-je vous appeler ?")
	uname = takeCommand()
	speak("Bienvenue ")
	speak(uname)
	columns = shutil.get_terminal_size().columns
	
	print("#####################".center(columns))
	print("Bienvenue ", uname.center(columns))
	print("#####################".center(columns))
	
	speak("Comment puis-je vous aider ?")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Écoute...")
		r.pause_threshold = 1
		audio = r.listen(source)

	try:
		print("Reconnaissance...")
		query = r.recognize_google(audio, language ='fr-FR')
		print(f"Vous avez dis: {query}\n")

	except Exception as e:
		print(e)
		print("Impossible de reconnaître votre voix.")
		return "Aucun"
	
	return query

def sendEmail(to, content):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	
	# Activer la faible sécurité dans gmail
	server.login('votre identifiant de messagerie', 'votre mot de passe de messagerie')
	server.sendmail('votre identifiant de messagerie', to, content)
	server.close()


if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# Cette fonction nettoiera toutes
	# commande avant l'exécution de ce fichier python
	clear()
	wishMe()
	username()
	
	while True:
		
		query = takeCommand().lower()
		
		# Toutes les commandes prononcées par l'utilisateur seront
		# stocké ici dans 'requête' et sera
		# converti en minuscules pour facilement
		# reconnaitre les commandements
		if 'wikipedia' in query:
			speak('Recherche Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences = 3)
			speak("Selon Wikipedia")
			print(results)
			speak(results)

		elif 'ouvre youtube' in query:
			speak("Allons sur Youtube\n")
			webbrowser.open("youtube.com")

		elif 'ouvre google' in query:
			speak("Allons sur Google\n")
			webbrowser.open("google.com")

		elif 'ouvre stackoverflow' in query:
			speak("Allons sur Stack Over flow.Happy coding")
			webbrowser.open("stackoverflow.com")

		elif ' musique' in query :
			speak("Go mettre de la musique")
			music_dir = "C:\\Users\\azfyxi\\Music"
			songs = os.listdir(music_dir)
			print(songs)
			random = os.startfile(os.path.join(music_dir, songs[1]))

		elif 'quelle heure est-il' in query:
			strTime = datetime.datetime.now().strftime("% H:% M:% S")
			speak(f"Il est {strTime}")

		elif 'ouvre chrome' in query:
			codePath = r"C:\\Users\\azfyxi\\AppData\\Local\\Programs\\chrome\\chrome.exe"
			os.startfile(codePath)

		elif 'envoyer à azfyxi' in query:
			try:
				speak("Qu'est-ce que je devrais dire?")
				content = takeCommand()
				to = "Adresse e-mail du destinataire"
				sendEmail(to, content)
				speak("L'email a été envoyé !")
			except Exception as e:
				print(e)
				speak("Je n'arrive pas à envoyer cet email")

		elif 'envoie un mail' in query:
			try:
				speak("Qu'est-ce que je devrais dire ?")
				content = takeCommand()
				speak("à qui dois-je envoyer")
				to = input()
				sendEmail(to, content)
				speak("L'email a été envoyé !")
			except Exception as e:
				print(e)
				speak("Je n'arrive pas à envoyer cet email")

		elif 'Comment ça va' in query:
			speak("Je vais bien, merci")
			speak("Comment allez-vous ?")

		elif 'bien' in query :
			speak("C'est bon de savoir que vous allez bien")

		elif "changer mon nom en" in query:
			query = query.replace("changer mon nom en", "")
			assname = query

		elif "changer de nom" in query:
			speak("Comment voudriez-vous m'appeler")
			assname = takeCommand()
			speak("Merci de m'avoir nommé")

		elif "comment tu t'appelles" in query:
			speak("Mes amis m'appellent")
			speak(assname)
			print("Mes amis m'appellent", assname)

		elif 'stop' in query:
			speak("Merci de m'avoir accordé de votre temps")
			exit()

		elif "qui t'a fait" in query or "qui t'a créé" in query:
			speak("J'ai été créé par azfyxi.")
			
		elif 'blague' in query:
			speak(pyjokes.get_joke())
			
		elif "calcule" in query:
			
			app_id = "Wolframalpha api id"
			client = wolframalpha.Client(app_id)
			indx = query.lower().split().index('calculate')
			query = query.split()[indx + 1:]
			res = client.query(' '.join(query))
			answer = next(res.results).text
			print("La réponse est " + answer)
			speak("La réponse est " + answer)

		elif 'cherche' in query or 'joue' in query:
			
			query = query.replace("cherche", "")
			query = query.replace("joue", "")		
			webbrowser.open(query)

		elif "qui suis-je" in query:
			speak("Si vous parlez alors certainement un humain.")

		elif 'présentation du projet' in query:
			speak("ouverture de la présentation de Sira")
			power = r"C:\\Users\\azfyxi\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
			os.startfile(power)

		elif "qui es-tu" in query:
			speak("Je suis votre assistant virtuel")

		elif 'changement de fond' in query:
			ctypes.windll.user32.SystemParametersInfoW(20,
													0,
													"Emplacement du fond",
													0)
			speak("L'arrière-plan a été modifié avec succès")

		elif 'ouvre bluestack' in query:
			appli = r"C:\\ProgramData\\BlueStacks\\Client\\Bluestacks.exe"
			os.startfile(appli)

		elif 'nouvelles' in query:
			
			try:
				jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
				data = json.load(jsonObj)
				i = 1
				
				speak('voici quelques nouvelles')
				print('''=============== TIMES OF INDIA ============'''+ '\n')
				
				for item in data['articles']:
					
					print(str(i) + '. ' + item['Titre'] + '\n')
					print(item['description'] + '\n')
					speak(str(i) + '. ' + item['Titre'] + '\n')
					i += 1
			except Exception as e:
				
				print(str(e))

		
		elif 'verrouille windows' in query:
				speak("verrouillage de l'appareil")
				ctypes.windll.user32.LockWorkStation()

		elif 'arrête le système' in query:
				speak("Attendez une seconde ! Votre système est sur le point de s'arrêter")
				subprocess.call('shutdown / p /f')
				
		elif 'vide la poubelle' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Corbeille Recyclé")

		elif "n'écoute plus" in query or "arrête d'écouter" in query:
			speak("pendant combien de temps vous voulez empêcher Sira d'écouter les commandes")
			a = int(takeCommand())
			time.sleep(a)
			print(a)

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl / maps / place/" + location + "")

		elif "redémarre" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "met en veille" in query or "dors" in query:
			speak("mise en veille en cours")
			subprocess.call("shutdown / h")

		elif "Déconnecte la session" in query :
			speak("Assurez-vous que toutes les applications sont fermées avant de vous déconnecter")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "écris une note" in query:
			speak("Que dois-je écrire")
			note = takeCommand()
			file = open('sira.txt', 'w')
			speak("Dois-je inclure la date et l'heure")
			snfm = takeCommand()
			if 'oui' in snfm or 'bien sur' in snfm:
				strTime = datetime.datetime.now().strftime("% H:% M:% S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
			else:
				file.write(note)
		
		elif "affiche la note" in query:
			speak("Affichage des notes en cours")
			file = open("sira.txt", "r")
			print(file.read())
			speak(file.read(6))

		elif "météo" in query:
			
			# Google Ouvrir le site Web de la météo
			# pour obtenir l'API de la météo ouverte
			api_key = "Api key"
			base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
			speak(" Nom de Ville ")
			print("Nom de Ville : ")
			city_name = takeCommand()
			complete_url = base_url + "appid =" + api_key + "&q =" + city_name
			response = requests.get(complete_url)
			x = response.json()
			
			if x["cod"] != "404":
				y = x["main"]
				current_temperature = y["temp"]
				current_pressure = y["pressure"]
				current_humidiy = y["humidity"]
				z = x["weather"]
				weather_description = z[0]["description"]
				print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
			
			else:
				speak(" Ville introuvable ")
			
		elif "envoie un message " in query:
				# On doit créer un compte sur Twilio pour utiliser cette commande
				account_sid = 'Account Sid key'
				auth_token = 'Auth token'
				client = Client(account_sid, auth_token)

				message = client.messages \
								.create(
									body = takeCommand(),
									from_='Sender No',
									to ='Receiver No'
								)

				print(message.sid)

		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Bonjour" in query:
			speak("Un chaleureux" +query)
			speak("Comment allez-vous")
			speak(assname)

		elif "Comment ça va" in query:
			speak("Je vais bien, je suis contente d'être là")


		elif "qu'est-ce que" in query or "qui est" in query:
			
			# Utiliser la même clé API
			# qui à été généré précédemment
			client = wolframalpha.Client("API_ID")
			res = client.query(query)
			
			try:
				print (next(res.results).text)
				speak (next(res.results).text)
			except StopIteration:
				print ("Aucun résultat")

		# elif "" in query:
			# Pour ajouter plus de commandes
