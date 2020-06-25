import os 
import sys
import traceback
import os.path
import datetime

from string import digits
import subprocess

from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QLabel, QLineEdit, QWidget, QPushButton, QApplication, QHBoxLayout, QVBoxLayout
from PyQt5.QtCore import QObject, pyqtSignal, QRunnable, pyqtSlot, QThreadPool

# Signals and worker for multhreading
class WorkerSignals(QObject):

	start = pyqtSignal()
	finished = pyqtSignal()
	error = pyqtSignal(tuple)

class Worker(QRunnable):

	def __init__(self, fn, *args, **kwargs):

		super(Worker, self).__init__()

		# Store constructor arguments (re-used for processing)
		self.fn = fn
		self.args = args
		self.kwargs = kwargs
		self.signals = WorkerSignals()    
 

	# Run the differents signals and their results
	@pyqtSlot()
	def run(self):

		try:
			self.signals.start.emit()
			result = self.fn(*self.args, **self.kwargs)
		except:
			traceback.print_exc()
			exctype, value = sys.exc_info()[:2]
			self.signals.error.emit((exctype, value, traceback.format_exc()))
		finally:
			self.signals.finished.emit()  

# Main class for the script
class Motivateur(QWidget):

	def __init__(self):
		super().__init__()

		self.initUI()

	def initUI(self):

		# Set the root path
		self.path = os.getcwd()

		# Set the labels 
		label_name = QLabel("Nom et Prenom:")
		label_adress = QLabel("Adresse:")
		label_city = QLabel("Ville et code postal:")
		label_mail = QLabel("Mail:")
		label_number = QLabel("Telephone")
		label_job = QLabel("Intitule du poste:")
		label_diploma = QLabel("Diplôme")
		label_univ = QLabel("Ecole")
		label_contact = QLabel("Nom du destinataire (Avec intitulé)")
		label_comp = QLabel("3 compétences aleatoires:")

		# Set the input field
		self.input_name = QLineEdit(self)
		self.input_adress = QLineEdit(self)
		self.input_city = QLineEdit(self)
		self.input_mail = QLineEdit(self)
		self.input_number = QLineEdit(self)
		self.input_job = QLineEdit(self)
		self.input_diploma = QLineEdit(self)
		self.input_univ = QLineEdit(self)
		self.input_contact = QLineEdit(self)
		self.input_comp1 = QLineEdit(self)
		self.input_comp2 = QLineEdit(self)
		self.input_comp3 = QLineEdit(self)

		# Set the placeholders
		self.input_name.setPlaceholderText("Clement Sedack")
		self.input_adress.setPlaceholderText("30 Rue Patrick Balkany")
		self.input_city.setPlaceholderText("92300 Levallois-Perret")
		self.input_mail.setPlaceholderText("csedack@gmail.com")
		self.input_number.setPlaceholderText("123-456-789")
		self.input_job.setPlaceholderText("Deputé")
		self.input_diploma.setPlaceholderText("Master en étude du vide")
		self.input_univ.setPlaceholderText("L'université de Issou") 
		self.input_contact.setPlaceholderText("Laisser vide si pas de nom")
		self.input_comp1.setPlaceholderText("Dépythonage de puits")
		self.input_comp2.setPlaceholderText("Survie")
		self.input_comp3.setPlaceholderText("Rendez-vous CAF")

		# Set the button
		self.btn_launch = QPushButton("Générer")

		# Set the threadpool
		self.threadpool = QThreadPool()
		self.threadpool.setMaxThreadCount(5)

		# Set the competences box
		box_comp = QHBoxLayout()
		box_comp.addWidget(self.input_comp1)
		box_comp.addWidget(self.input_comp2)
		box_comp.addWidget(self.input_comp3)

		# Set main windows layout
		box_main = QVBoxLayout()
		box_main.addWidget(label_name)
		box_main.addWidget(self.input_name)
		box_main.addWidget(label_adress)
		box_main.addWidget(self.input_adress)
		box_main.addWidget(label_city)
		box_main.addWidget(self.input_city)
		box_main.addWidget(label_mail)
		box_main.addWidget(self.input_mail)
		box_main.addWidget(label_number)
		box_main.addWidget(self.input_number)
		box_main.addWidget(label_job)
		box_main.addWidget(self.input_job)
		box_main.addWidget(label_diploma)
		box_main.addWidget(self.input_diploma)
		box_main.addWidget(label_univ)
		box_main.addWidget(self.input_univ)
		box_main.addWidget(label_comp)
		box_main.addLayout(box_comp)
		box_main.addWidget(label_contact)
		box_main.addWidget(self.input_contact)
		box_main.addWidget(self.btn_launch)

		# Setting the main windows
		self.setLayout(box_main)
		self.setGeometry(300, 300,450, 50)
		self.setWindowTitle('Motivateur')
		self.setWindowIcon(QIcon('web.png'))        
		self.show()

		# Connect button to the launcher
		self.btn_launch.clicked.connect(self.launcher)
	
	# Multi thread caller for the script
	def launcher(self):

		worker = Worker(self.generator)
		self.threadpool.start(worker)

		worker.signals.start.connect(self.function_start)
		worker.signals.finished.connect(self.function_end)

	# main program function
	def generator(self):

		# Set the user inputs vars
		name = str(self.input_name.text())
		adress = str(self.input_adress.text())
		city = str(self.input_city.text())
		mail = str(self.input_mail.text())
		number = str(self.input_number.text())
		job = "Objet: Candidature au poste de \"{}\"".format(str(self.input_job.text()))
		diploma = str(self.input_diploma.text())
		univ = str(self.input_univ.text())
		comp1 = str(self.input_comp1.text())
		comp2 = str(self.input_comp2.text())
		comp3 = str(self.input_comp3.text())

		# Set the place and date
		remove_digits = str.maketrans('', '', digits)
		place = city.translate(remove_digits)
		place = place.strip()
		today = datetime.date.today()
		dateplace =  "À " + place + " le " + str(today.day) + "/" + str(today.month) + "/" + str(today.year)
		contact = "Madame, Monsieur,"
		end = contact
		# If there is text in the contact field
		if len(str(self.input_contact.text())) != 0 : 
			contact = "à l'attention de " + str(self.input_contact.text()) + ","
			end = str(self.input_contact.text()) + ","

		# Setting the path for the output
		letter = []
		outfilename = "Lettre.pdf"
		outfilepath = os.path.join( self.path, outfilename )

		# Setting the pdf document and styles 
		doc = SimpleDocTemplate(outfilepath, rightMargin=2*cm,leftMargin=2*cm,topMargin=2*cm,bottomMargin=2*cm)
		styles=getSampleStyleSheet()
		styles.add(ParagraphStyle(name='Right', alignment=TA_RIGHT))
		styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER))
		
		# Write the pdf with differents strings and inputs 
		ptext = '<font size="12">{}</font>'.format(name)
		letter.append(Paragraph(ptext, styles["Normal"]))  
		ptext = '<font size="12">{}</font>'.format(adress)
		letter.append(Paragraph(ptext, styles["Normal"]))
		ptext = '<font size="12">{}</font>'.format(city)
		letter.append(Paragraph(ptext, styles["Normal"]))
		ptext = '<font size="12">{}</font>'.format(mail)
		letter.append(Paragraph(ptext, styles["Normal"]))
		ptext = '<font size="12">{}</font>'.format(number)
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 50))
		ptext = '<font size="12">{}</font>'.format(dateplace)
		letter.append(Paragraph(ptext, styles["Right"])) 
		letter.append(Spacer(1, 80))
		ptext = '<font size="12">{}</font>'.format(job)
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 50))
		ptext = '<font size="12">{}</font>'.format(contact)
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">Ayant récemment obtenu mon diplôme de {} à {}, je suis désormais à la recherche d\'un emploi.</font>'.format(diploma, univ)
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">J\'ai, durant mon cursus scolaire et professionnel, pu acquérir de nombreuses compétences en {} mais aussi en {} ou encore en {}. Ces différentes expériences m\'ont permis d\'obtenir mes premiers savoirs et je pense désormais être en mesure de pouvoir candidater pour le poste que vous proposez aujourd\'hui.</font>'.format(comp1, comp2, comp3)
		letter.append(Paragraph(ptext, styles["Normal"])) 
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">Comme vous avez également pu le remarquer durant la lecture de mon curriculum vitae j\'ai aussi pu développer durant cette période plusieurs compétences annexes  sur mon temps personnel, qui, je pense, peuvent entrer en complémentarité avec les qualités requise pour occuper ce poste.</font>'
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">De plus, je trouve le fait de travailler pour une organisation d-envergure comme la vôtre peut être extrêmement enrichissant tant professionnellement que personnellement.</font>'
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">Appliqué, honnête et sociable, je souhaite occuper ce poste avec tout le sérieux et l\'enthousiasme dont je fais déjà preuve dans la poursuite de mes études. Mes capacités d’adaptation me permettent de m’intégrer très rapidement au sein d’une équipe de travail.</font>'
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">Je reste à votre disposition pour toute information complémentaire, ou pour vous rencontrer lors d’un entretien.</font>'
		letter.append(Paragraph(ptext, styles["Normal"]))  
		letter.append(Spacer(1, 15))
		ptext = '<font size="12">Veuillez agréer, {} l’expression de mes sincères salutations.</font>'.format(end)
		letter.append(Paragraph(ptext, styles["Normal"]))
		letter.append(Spacer(1, 30))
		ptext = '<font size="12">{}</font>'.format(name)
		letter.append(Paragraph(ptext, styles["Center"]))  

		# Build the pdf
		doc.build(letter)

		if os.name != 'nt':

			subprocess.call(["xdg-open", outfilepath])
		else :
			os.startfile(outfilepath)
	# When script end
	def function_end(self):

		# Re-enable the button
		self.btn_launch.setEnabled(True)

	# When script start
	def function_start(self):

			# Prevent the user to re-run the scrapper by disabling the button
			self.btn_launch.setEnabled(False)

if __name__ == '__main__':

	app = QApplication(sys.argv)
	ex = Motivateur()
	sys.exit(app.exec_())