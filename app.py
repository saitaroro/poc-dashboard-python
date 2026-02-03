import os
import pandas as pd
import matplotlib
matplotlib.use('Agg') # Pour éviter les erreurs d'interface graphique dans Docker
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, send_file, redirect, url_for
from pptx import Presentation
from pptx.util import Inches
from email.message import EmailMessage
import io

app = Flask(__name__)

# Configuration des dossiers
DATA_FILE = 'data/data_source.csv'
OUTPUT_DIR = 'output'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- 1. Génération de fausses données pour le POC ---
def create_mock_data():
    if not os.path.exists(DATA_FILE):
        os.makedirs('data', exist_ok=True)
        data = {
            'Bureau': ['Paris', 'Lyon', 'Marseille', 'Paris', 'Lyon'] * 24,
            'Profil': ['Junior', 'Senior', 'Manager', 'Directeur', 'Junior'] * 24,
            'Mois': ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin'] * 20,
            'Rdv': [15, 22, 10, 5, 18] * 24
        }
        df = pd.DataFrame(data)
        # On ajuste pour avoir 120 lignes pour l'exemple
        df = df.iloc[:120] 
        df.to_csv(DATA_FILE, index=False)

# --- 2. Fonctions de Calcul et Graphiques ---
def process_data():
    df = pd.read_csv(DATA_FILE)
    
    # KPI 1: Total utilisateurs (simulé par nombre de lignes uniques ici pour l'exemple)
    total_users = len(df)
    
    # KPI 2: Classement des bureaux
    top_bureaux = df.groupby('Bureau')['Rdv'].sum().sort_values(ascending=False)
    
    # KPI 3: Graphique par mois
    monthly_data = df.groupby('Mois')['Rdv'].sum()
    
    # Création du graphique
    plt.figure(figsize=(10, 6))
    monthly_data.plot(kind='bar', color='skyblue')
    plt.title('Nombre de RDV par Mois en 2025')
    plt.xlabel('Mois')
    plt.ylabel('Nombre de RDV')
    plt.tight_layout()
    chart_path = os.path.join(OUTPUT_DIR, 'chart_mois.png')
    plt.savefig(chart_path)
    plt.close()
    
    return total_users, top_bureaux, chart_path

# --- 3. Génération du PowerPoint ---
def generate_pptx(total_users, top_bureaux, chart_path):
    prs = Presentation()
    
    # Slide 1: Titre
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Rapport Mensuel 2025"
    subtitle.text = "Analyse des Rendez-vous par Bureau"
    
    # Slide 2: Données et Graphique
    slide_layout = prs.slide_layouts[1] # Titre et contenu
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Vue d'ensemble"
    
    # Ajouter le texte des stats
    text_frame = slide.shapes.placeholders[1].text_frame
    text_frame.text = f"Total Utilisateurs traités : {total_users}"
    p = text_frame.add_paragraph()
    p.text = f"Meilleur Bureau : {top_bureaux.index[0]} ({top_bureaux.iloc[0]} rdv)"
    
    # Ajouter l'image du graphique
    slide.shapes.add_picture(chart_path, Inches(1), Inches(3.5), width=Inches(8))
    
    pptx_path = os.path.join(OUTPUT_DIR, 'Rapport_Final.pptx')
    prs.save(pptx_path)
    return pptx_path

# --- ROUTES FLASK (Le Site Web) ---

@app.route('/')
def index():
    create_mock_data() # S'assure que le fichier existe
    return render_template('index.html')

@app.route('/integrate', methods=['POST'])
def integrate():
    # Lance les calculs
    total, top, chart = process_data()
    # Génère le PPT
    generate_pptx(total, top, chart)
    return redirect(url_for('validation_page'))

@app.route('/validate')
def validation_page():
    return render_template('validate.html')

@app.route('/test-download')
def test_download():
    # Permet de télécharger le PPT pour vérifier
    return send_file(os.path.join(OUTPUT_DIR, 'Rapport_Final.pptx'), as_attachment=True)

@app.route('/generate-email', methods=['POST'])
def generate_email():
    # Récupération des données du formulaire
    email_to = request.form.get('email_to')
    email_cc = request.form.get('email_cc')
    subject = request.form.get('subject')
    
    # Création de l'objet Email
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = "monapp@entreprise.com" # Juste pour le format
    msg['To'] = email_to
    msg['Cc'] = email_cc
    msg.set_content("Bonjour,\n\nVeuillez trouver ci-joint le rapport mensuel des rendez-vous.\n\nCordialement.")
    
    # Attacher le PowerPoint
    pptx_path = os.path.join(OUTPUT_DIR, 'Rapport_Final.pptx')
    with open(pptx_path, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.presentationml.presentation', filename="Rapport_2025.pptx")
    
    # Sauvegarder en .eml
    eml_path = os.path.join(OUTPUT_DIR, 'brouillon_outlook.eml')
    with open(eml_path, 'wb') as f:
        f.write(msg.as_bytes())
        
    return send_file(eml_path, as_attachment=True, download_name="Ouvrir_dans_Outlook.eml")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)