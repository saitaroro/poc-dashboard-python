"""Module de prévision simple.

Contient une fonction professionnelle de vente pour entraîner un modèle
sur la série temporelle hebdomadaire et produire des prédictions pour les
6 mois à venir. Utilise une méthode dissage exponentiel triple (Holt-Winters)
via statsmodels.
"""

from __future__ import annotations

import os
from typing import Dict

import matplotlib.pyplot as plt
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing


def train_and_forecast(
    df: pd.DataFrame,
    output_dir: str,
    last_weeks: int = 26,
    horizon_weeks: int = 26,
) -> Dict[str, pd.Series]:
    """Entraîne un modèle à partir des rendez-vous et prédit les prochaines semaines.

    Le DataFrame d'entrée doit contenir les colonnes `date_rdv` et `nb_rdv`.
    La fonction:

    * génère la série hebdomadaire des volumes,
    * entraîne un modèle ``ExponentialSmoothing`` avec saisonnalité hebdomadaire,
    * prédit ``horizon_weeks`` semaines dans le futur,
    * retourne les ``last_weeks`` valeurs réelles précédentes ainsi que la
      prédiction et enregistre un graphique dans ``output_dir``.

    Remarque: ``statsmodels`` est une dépendance supplémentaire (voir
    ``requirements.txt``).
    """
    # resample hebdomadaire (dimanche par défaut)
    series = (
        df.set_index('date_rdv')['nb_rdv']
        .resample('W')
        .sum()
        .asfreq('W')
        .fillna(0)
    )

    if len(series) < last_weeks:
        raise ValueError('Pas assez de données pour calculer la prévision')

    train = series.iloc[-last_weeks:]

    model_path = os.path.join(output_dir, 'forecast_model.pkl')
    if os.path.exists(model_path):
        # charger modèle existant pour gagner du temps
        import pickle

        with open(model_path, 'rb') as f:
            fit = pickle.load(f)
    else:
        # modèle simple additive avec saisonnalité hebdomadaire
        model = ExponentialSmoothing(
            train,
            trend='add',
            seasonal='add',
            seasonal_periods=52,
        )
        fit = model.fit(optimized=True)
        # sauvegarder pour réutilisation
        import pickle

        with open(model_path, 'wb') as f:
            pickle.dump(fit, f)

    forecast = fit.forecast(horizon_weeks)

    # tracé
    plt.figure(figsize=(10, 6))
    plt.plot(train.index, train.values, label='réel (6 derniers mois)')
    plt.plot(forecast.index, forecast.values, label='prévision 6 mois', linestyle='--')
    plt.title('Prévision hebdomadaire des rendez-vous')
    plt.xlabel('Date')
    plt.ylabel('Nombre de rendez-vous')
    plt.legend()
    plt.tight_layout()

    chart_path = os.path.join(output_dir, 'forecast.png')
    plt.savefig(chart_path)
    plt.close()

    return {
        'actual': train,
        'forecast': forecast,
        'chart': chart_path,
    }
