# Utiliser une image Python officielle légère
FROM python:3.9-slim

# Définir le dossier de travail dans le conteneur
WORKDIR /app

# Copier les fichiers requis
COPY requirements.txt .

# Installer les dépendances
RUN pip install --no-cache-dir -r requirements.txt

# Copier tout le code dans le conteneur
COPY . .

# Exposer le port 5000 (celui de Flask)
EXPOSE 5000

# Commande de démarrage
CMD ["python", "app.py"]
