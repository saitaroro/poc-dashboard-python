import logging
import subprocess
import sys
from pathlib import Path


logger = logging.getLogger('run_local')
handler = logging.StreamHandler(sys.stdout)
fmt = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
handler.setFormatter(fmt)
logger.addHandler(handler)
logger.setLevel(logging.INFO)


ROOT = Path(__file__).parent


def main():
    logger.info('Lancement du helper de démarrage local')

    # Installer les dépendances (silencieux sauf erreur)
    req = ROOT / 'requirements.txt'
    if req.exists():
        logger.info('Installation des dépendances depuis %s', req)
        r = subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', str(req)])
        if r.returncode != 0:
            logger.error('Erreur lors de l installation des dépendances')
            sys.exit(r.returncode)
    else:
        logger.warning('requirements.txt introuvable, saut de l installation')

    # Démarrer l'application
    logger.info('Démarrage de l application via %s', ROOT / 'app.py')
    try:
        subprocess.run([sys.executable, str(ROOT / 'app.py')], check=True)
    except subprocess.CalledProcessError as e:
        logger.error('L application s est terminée avec le code %s', e.returncode)
        sys.exit(e.returncode)


if __name__ == '__main__':
    main()
