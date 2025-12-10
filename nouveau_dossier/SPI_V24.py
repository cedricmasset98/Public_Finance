import sys
import os
import glob
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QComboBox, QPushButton, QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget,
    QListWidget, QListWidgetItem, QAbstractItemView, QDateEdit, QCheckBox, QSpinBox
)
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'Arial'  # Police par défaut pour tous les graphiques
plt.rcParams['font.size'] = 14           # Taille générale
plt.rcParams['axes.titlesize'] = 16      # Titres des axes
plt.rcParams['axes.labelsize'] = 15      # Labels des axes
plt.rcParams['xtick.labelsize'] = 13     # Graduation axe X
plt.rcParams['ytick.labelsize'] = 13     # Graduation axe Y
plt.rcParams['legend.fontsize'] = 13     # Légende
plt.rcParams['figure.titlesize'] = 18    # Titre de la figure
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf

    # --- Correction centrage graphique rendement ---
def recenter_return_plot(self):
    self.figure_return.tight_layout()
    self.canvas_return.draw()
    self.canvas_return.flush_events()

def on_tab_changed(self, index):
        # Si l'onglet courant est "Rendements"
        if self.tabs.tabText(index).startswith("📊 Rendements"):
            self.recenter_return_plot()

class MainWindow(QWidget):
    def recenter_return_plot(self):
        self.figure_return.tight_layout()
        self.canvas_return.draw()
        self.canvas_return.flush_events()

    def recenter_qq_plot(self):
        self.figure_qq.tight_layout()
        self.canvas_qq.draw()
        self.canvas_qq.flush_events()

    def recenter_scatter_plot(self):
        self.figure_scatter.tight_layout()
        self.canvas_scatter.draw()
        self.canvas_scatter.flush_events()

    def on_tab_changed(self, index):
        tab_text = self.tabs.tabText(index)
        if tab_text.startswith("📊 Rendements"):
            self.recenter_return_plot()
        elif tab_text.startswith("📐 QQ-Plot"):
            self.recenter_qq_plot()
        elif tab_text.startswith("🔵 Nuage de points"):
            self.recenter_scatter_plot()

    def get_asset_names(self):
        # Liste tous les fichiers .xlsx du dossier courant et extrait les noms (sans extension)
        excel_files = glob.glob(os.path.join(os.path.dirname(__file__), '*.xlsx'))
        asset_names = [os.path.splitext(os.path.basename(f))[0] for f in excel_files if not os.path.basename(f).startswith('~$')]
        return asset_names

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cours de bourse SPI")
        self.setGeometry(50, 50, 1400, 900)
        main_layout = QVBoxLayout(self)



        # --- Détection dynamique des actifs (Excel) ---
        self.asset_names = self.get_asset_names()

        # --- Toolbar en haut ---
        top_layout = QHBoxLayout()
        self.asset_select = QComboBox()
        self.asset_select.addItems(self.asset_names)
        top_layout.addWidget(self.asset_select)
        self.asset_select.currentIndexChanged.connect(self.update_date_range_from_asset)

        self.date_start = QDateEdit()
        self.date_start.setCalendarPopup(True)
        self.date_start.setDisplayFormat('dd.MM.yyyy')

        # Préremplir avec la première date des données
        import os
        first_date = None
        try:
            if os.path.exists('SPI.xlsx'):
                df_dates = pd.read_excel('SPI.xlsx', engine='openpyxl', decimal=',')
                if 'Date' in df_dates.columns:
                    df_dates['Date'] = pd.to_datetime(df_dates['Date'], format='%d.%m.%Y', errors='coerce')
                    df_dates = df_dates.dropna(subset=['Date'])
                    if not df_dates.empty:
                        first_date = df_dates['Date'].min()
        except Exception as e:
            first_date = None
        # Les dates seront mises à jour dynamiquement selon l’actif
        self.date_start.setFixedWidth(120)  # largeur augmentée
        top_layout.addWidget(QLabel("Début :"))
        top_layout.addWidget(self.date_start)

        self.date_end = QDateEdit()
        self.date_end.setCalendarPopup(True)
        self.date_end.setDisplayFormat('dd.MM.yyyy')
        self.date_end.setFixedWidth(120)  # largeur augmentée
        top_layout.addWidget(QLabel("Fin :"))
        top_layout.addWidget(self.date_end)

        # Met à jour les dates au démarrage
        self.update_date_range_from_asset()


        self.all_series_checkbox = QCheckBox("Toute la série")
        self.all_series_checkbox.setChecked(False)  # décochée par défaut
        self.all_series_checkbox.stateChanged.connect(self.toggle_date_edits)
        top_layout.addWidget(self.all_series_checkbox)

        # Choix de la fréquence
        self.freq_select = QComboBox()
        self.freq_select.addItems(["Journalier", "Hebdomadaire", "Mensuelle", "Annuelle"])
        self.freq_select.setCurrentIndex(0)
        top_layout.addWidget(QLabel("Fréquence :"))
        top_layout.addWidget(self.freq_select)

        # Choix du type de rendement
        self.return_type_select = QComboBox()
        self.return_type_select.addItems(["Simple", "Logarithmique"])
        self.return_type_select.setCurrentIndex(0)
        top_layout.addWidget(QLabel("Type de rendement :"))
        top_layout.addWidget(self.return_type_select)

        top_layout.addStretch()
        self.btn = QPushButton("Afficher le graphique")
        self.btn.setStyleSheet("""
QPushButton {
    background-color: #4F8EF7;
    color: white;
    border-radius: 10px;
    padding: 10px 24px;
    font-size: 16px;
    font-weight: bold;
    margin-left: 12px;
}
QPushButton:hover {
    background-color: #2566d8;
}
QPushButton:pressed {
    background-color: #174a8c;
    border: 1px solid #174a8c;
}
""")
        self.btn.clicked.connect(self.update_graphs)
        top_layout.addWidget(self.btn)

        main_layout.addLayout(top_layout)

        # --- Tabs ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs, stretch=1)

        # --- Onglet graphique ---
        self.tab_price = QWidget()
        self.tabs.addTab(self.tab_price, "📈 Cours")
        tab_price_layout = QVBoxLayout(self.tab_price)
        tab_price_layout.setContentsMargins(10, 10, 10, 10)
        tab_price_layout.setSpacing(10)

        self.figure, (self.ax_price, self.ax_drawdown) = plt.subplots(nrows=2, ncols=1, figsize=(12, 8), gridspec_kw={'height_ratios': [3, 1]}, sharex=True)
        self.canvas = FigureCanvas(self.figure)
        tab_price_layout.addWidget(self.canvas, stretch=1)

        # --- Onglet rendements ---
        self.tab_return = QWidget()
        self.tabs.addTab(self.tab_return, "📊 Rendements")
        tab_return_layout = QVBoxLayout(self.tab_return)
        tab_return_layout.setContentsMargins(10, 10, 10, 10)
        tab_return_layout.setSpacing(10)

        # Deux subplots : [0] série temporelle, [1] histogramme
        self.figure_return, (self.ax_return, self.ax_hist) = plt.subplots(nrows=2, ncols=1, figsize=(12, 10), gridspec_kw={'height_ratios': [2, 1]})
        self.canvas_return = FigureCanvas(self.figure_return)
        tab_return_layout.addWidget(self.canvas_return, stretch=1)

        # Connexion centrage automatique de l'onglet rendement
        self.tabs.currentChanged.connect(self.on_tab_changed)



        # --- Onglet Nuage de points ---
        self.tab_scatter = QWidget()
        self.tabs.addTab(self.tab_scatter, "🔵 Nuage de points")
        scatter_layout = QVBoxLayout(self.tab_scatter)
        scatter_layout.setContentsMargins(10, 10, 10, 10)
        scatter_layout.setSpacing(10)
        self.figure_scatter, self.ax_scatter = plt.subplots()
        self.canvas_scatter = FigureCanvas(self.figure_scatter)
        scatter_layout.addWidget(self.canvas_scatter, stretch=1)

        # --- Onglet Moments glissants ---
        self.tab_moments = QWidget()
        self.tabs.addTab(self.tab_moments, "📈 Moments glissants")
        tab_moments_layout = QVBoxLayout(self.tab_moments)
        tab_moments_layout.setContentsMargins(10, 10, 10, 10)
        tab_moments_layout.setSpacing(10)
        # Contrôles : sélecteur de fenêtre et bouton
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(QLabel("Taille de la fenêtre glissante (jours) :"))
        self.rolling_window_spin = QSpinBox()
        self.rolling_window_spin.setMinimum(5)
        self.rolling_window_spin.setMaximum(2**31 - 1)
        self.rolling_window_spin.setValue(252)
        controls_layout.addWidget(self.rolling_window_spin)
        self.btn_calc_rolling = QPushButton("Calculer")
        controls_layout.addWidget(self.btn_calc_rolling)
        controls_layout.addStretch()
        tab_moments_layout.addLayout(controls_layout)
        # Figure matplotlib 2x2
        import matplotlib.gridspec as gridspec
        self.figure_moments = plt.figure(figsize=(12, 10))
        gs_mom = gridspec.GridSpec(2, 2)
        self.ax_mean = self.figure_moments.add_subplot(gs_mom[0, 0])
        self.ax_vol = self.figure_moments.add_subplot(gs_mom[0, 1])
        self.ax_skew = self.figure_moments.add_subplot(gs_mom[1, 0])
        self.ax_kurt = self.figure_moments.add_subplot(gs_mom[1, 1])
        self.canvas_moments = FigureCanvas(self.figure_moments)
        tab_moments_layout.addWidget(self.canvas_moments, stretch=1)
        self.btn_calc_rolling.clicked.connect(self.plot_rolling_moments)

        # --- Onglet QQ-Plot ---
        self.tab_qq = QWidget()
        self.tabs.addTab(self.tab_qq, "📐 QQ-Plot")
        qq_layout = QVBoxLayout(self.tab_qq)
        qq_layout.setContentsMargins(10, 10, 10, 10)
        qq_layout.setSpacing(10)
        # Figure : QQ-Plot (2/3 sup), ACF & PACF côte à côte (1/3 inf)
        import matplotlib.gridspec as gridspec
        self.figure_qq = plt.figure(figsize=(12, 12))
        gs = gridspec.GridSpec(2, 2, height_ratios=[2, 1])
        self.ax_qq = self.figure_qq.add_subplot(gs[0, :])  # QQ-Plot sur toute la largeur
        self.ax_acf = self.figure_qq.add_subplot(gs[1, 0]) # ACF à gauche
        self.ax_pacf = self.figure_qq.add_subplot(gs[1, 1]) # PACF à droite
        self.canvas_qq = FigureCanvas(self.figure_qq)
        qq_layout.addWidget(self.canvas_qq, stretch=1)
        qq_layout.addStretch()

        # --- Onglet Statistiques ---
        self.tab_stats = QWidget()
        self.tabs.addTab(self.tab_stats, "📋 Statistiques")
        stats_layout = QVBoxLayout(self.tab_stats)
        stats_layout.setContentsMargins(10, 10, 10, 10)
        stats_layout.setSpacing(10)
        self.stats_table = QTableWidget(10, 2)
        self.stats_table.setHorizontalHeaderLabels(["Statistique", "Valeur"])
        self.stats_table.setVerticalHeaderLabels(["" for _ in range(10)])
        self.stats_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.stats_table.verticalHeader().setVisible(False)
        self.stats_table.setEditTriggers(QTableWidget.NoEditTriggers)
        stats_layout.addWidget(self.stats_table)

        # --- Onglet Comparative cumulative Return ---
        from PyQt5.QtWidgets import QListWidget, QListWidgetItem, QAbstractItemView
        self.tab_cumret = QWidget()
        self.tabs.addTab(self.tab_cumret, "📈 Comparative cumulative Return")
        cumret_layout = QVBoxLayout(self.tab_cumret)
        cumret_layout.setContentsMargins(10, 10, 10, 10)
        cumret_layout.setSpacing(10)
        # Sélection multiple des actifs (User Friendly)

        cumret_select_layout = QHBoxLayout()
        cumret_select_layout.addWidget(QLabel("Actifs :"))
        self.cumret_asset_list = QListWidget()
        self.cumret_asset_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.cumret_asset_list.setFixedWidth(180)
        # Remplir dynamiquement la liste avec cases à cocher
        for asset in self.asset_names:
            item = QListWidgetItem(asset)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            font = self.cumret_asset_list.font()
            font.setPointSize(13)
            item.setFont(font)
            self.cumret_asset_list.addItem(item)
        cumret_select_layout.addWidget(self.cumret_asset_list)
        # Boutons Tout sélectionner / Tout désélectionner
        btns_layout = QVBoxLayout()
        self.cumret_select_all_btn = QPushButton("Tout sélectionner")
        self.cumret_deselect_all_btn = QPushButton("Tout désélectionner")
        btns_layout.addWidget(self.cumret_select_all_btn)
        btns_layout.addWidget(self.cumret_deselect_all_btn)
        btns_layout.addStretch()
        cumret_select_layout.addLayout(btns_layout)
        # Label résumé sélection (désactivé)
        self.cumret_selected_label = QLabel("")
        cumret_select_layout.addWidget(self.cumret_selected_label)
        cumret_select_layout.addStretch()
        self.cumret_btn = QPushButton("Calculer")
        btns_layout.addWidget(self.cumret_btn)
        cumret_layout.addLayout(cumret_select_layout)
        # Figure matplotlib pour le comparatif
        self.figure_cumret, self.ax_cumret = plt.subplots(figsize=(12, 6))
        self.canvas_cumret = FigureCanvas(self.figure_cumret)
        cumret_layout.addWidget(self.canvas_cumret, stretch=1)
        self.cumret_btn.clicked.connect(self.plot_cumret)
        # Connecter les boutons
        self.cumret_select_all_btn.clicked.connect(self.cumret_select_all)
        self.cumret_deselect_all_btn.clicked.connect(self.cumret_deselect_all)
        self.cumret_asset_list.itemChanged.connect(self.update_cumret_selected_label)
        self.update_cumret_selected_label()

        # --- Onglet Rolling Return ---
        self.tab_rolling = QWidget()
        self.tabs.addTab(self.tab_rolling, "🔄 Rolling Return")
        rolling_layout = QVBoxLayout(self.tab_rolling)
        rolling_layout.setContentsMargins(10, 10, 10, 10)
        rolling_layout.setSpacing(10)
        rolling_top = QHBoxLayout()
        rolling_top.addWidget(QLabel("Fenêtre glissante :"))
        self.rolling_combo = QComboBox()
        self.rolling_combo.addItems(["1 an", "3 ans", "5 ans", "10 ans"])
        self.rolling_combo.setCurrentIndex(0)
        rolling_top.addWidget(self.rolling_combo)
        self.rolling_btn = QPushButton("Calculer")
        rolling_top.addWidget(self.rolling_btn)
        rolling_top.addStretch()
        rolling_layout.addLayout(rolling_top)
        # Figure matplotlib pour rolling return : 2 sous-graphiques (série + histogramme)
        self.figure_rolling, (self.ax_rolling, self.ax_rolling_hist) = plt.subplots(nrows=2, ncols=1, figsize=(12, 10), gridspec_kw={'height_ratios': [2, 1]})
        self.canvas_rolling = FigureCanvas(self.figure_rolling)
        rolling_layout.addWidget(self.canvas_rolling, stretch=1)
        # Connexion du bouton uniquement
        self.rolling_btn.clicked.connect(self.plot_rolling_return)



    def update_date_range_from_asset(self):
        from PyQt5.QtCore import QDate
        actif = self.asset_select.currentText()
        file_path = os.path.join(os.path.dirname(__file__), actif + '.xlsx')
        try:
            if os.path.exists(file_path):
                df_dates = pd.read_excel(file_path, engine='openpyxl', decimal=',')
                if 'Date' in df_dates.columns:
                    df_dates['Date'] = pd.to_datetime(df_dates['Date'], format='%d.%m.%Y', errors='coerce')
                    df_dates = df_dates.dropna(subset=['Date'])
                    if not df_dates.empty:
                        first_date = df_dates['Date'].min()
                        last_date = df_dates['Date'].max()
                        self.date_start.setDate(QDate(first_date.year, first_date.month, first_date.day))
                        self.date_end.setDate(QDate(last_date.year, last_date.month, last_date.day))
                        return
        except Exception as e:
            pass  # En cas d’erreur, ne rien changer
        # Si problème, valeurs par défaut
        from PyQt5.QtCore import QDate
        self.date_start.setDate(QDate.currentDate())
        self.date_end.setDate(QDate.currentDate())

    def toggle_date_edits(self):
        checked = self.all_series_checkbox.isChecked()
        self.date_start.setEnabled(not checked)
        self.date_end.setEnabled(not checked)
        # Suppression de la mise à jour automatique du graphique

    def plot_cumret(self):
        import numpy as np
        from datetime import datetime
        # Nouvelle sélection à partir des cases cochées
        assets = [self.cumret_asset_list.item(i).text() for i in range(self.cumret_asset_list.count()) if self.cumret_asset_list.item(i).checkState() == Qt.Checked]
        if not assets:
            QMessageBox.warning(self, "Aucun actif sélectionné", "Veuillez sélectionner au moins un actif.")
            return
        date_debut = self.date_start.date().toPyDate()
        date_fin = self.date_end.date().toPyDate()
        self.ax_cumret.clear()
        for asset in assets:
            file_path = os.path.join(os.path.dirname(__file__), asset + '.xlsx')
            try:
                df = pd.read_excel(file_path, engine='openpyxl', decimal=',')
                df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y')
                df = df.sort_values('Date')
                # Filtrage selon la période choisie
                if not self.all_series_checkbox.isChecked():
                    df = df[(df['Date'] >= pd.Timestamp(date_debut)) & (df['Date'] <= pd.Timestamp(date_fin))]
                if 'Close' in df.columns:
                    prices = df['Close']
                else:
                    prices = df.iloc[:, 1]  # 2e colonne si pas de nom
                returns = prices.pct_change().dropna()
                cumulative = (1 + returns).cumprod()
                if not cumulative.empty:
                    final_return = (cumulative.iloc[-1] - 1) * 100
                    label = f"{asset} ({final_return:.2f} %)"
                else:
                    label = asset
                self.ax_cumret.plot(df['Date'].iloc[1:], cumulative, label=label)
            except Exception as e:
                QMessageBox.warning(self, "Erreur", f"Erreur lors du chargement de {asset} : {str(e)}")
                continue
        self.ax_cumret.set_title("Comparative cumulative Return")
        self.ax_cumret.set_xlabel("Date")
        self.ax_cumret.set_ylabel("Rendement cumulé")
        self.ax_cumret.legend()
        self.ax_cumret.grid(True)
        self.figure_cumret.tight_layout()
        self.canvas_cumret.draw()

    def plot_rolling_return(self):
        import numpy as np
        from matplotlib.dates import DateFormatter
        import matplotlib.ticker as mticker
        actif = self.asset_select.currentText()
        file_path = os.path.join(os.path.dirname(__file__), actif + '.xlsx')
        title = f'Rendement glissant {actif}'
        if not os.path.exists(file_path):
            return
        df = pd.read_excel(file_path, engine='openpyxl', decimal=',')
        df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y')
        df = df.sort_values('Date')
        # Filtrage selon la période choisie
        if not self.all_series_checkbox.isChecked():
            date_debut = self.date_start.date().toPyDate()
            date_fin = self.date_end.date().toPyDate()
            df = df[(df['Date'] >= pd.Timestamp(date_debut)) & (df['Date'] <= pd.Timestamp(date_fin))]
        # Sinon, on garde toute la série
        # Type de rendement
        return_type = self.return_type_select.currentText()
        if return_type == "Simple":
            returns = df['Close'].pct_change()
        else:
            returns = np.log(df['Close'] / df['Close'].shift(1))
        # Fenêtre glissante
        window_label = self.rolling_combo.currentText()
        if window_label == "1 an":
            window = 252  # jours ouvrés
        elif window_label == "3 ans":
            window = 252 * 3
        elif window_label == "5 ans":
            window = 252 * 5
        elif window_label == "10 ans":
            window = 252 * 10
        else:
            window = 252
        rolling_return = (1 + returns).rolling(window).apply(np.prod, raw=True) - 1
        # Nettoyage des axes
        self.ax_rolling.clear()
        self.ax_rolling_hist.clear()
        # Tracé du rolling return
        self.ax_rolling.plot(df['Date'], rolling_return, color='#4F8EF7', label='Rolling Return')
        mean_val = rolling_return.mean()
        median_val = rolling_return.median()
        self.ax_rolling.axhline(mean_val, color='black', linestyle='-', linewidth=2, label=f'Moyenne ({mean_val:.2%})')
        self.ax_rolling.axhline(median_val, color='orange', linestyle='--', linewidth=2, label=f'Médiane ({median_val:.2%})')
        self.ax_rolling.set_title(f"{title} ({window_label})")
        self.ax_rolling.set_xlabel('Date')
        self.ax_rolling.set_ylabel('Rendement glissant')
        self.ax_rolling.grid(True)
        self.ax_rolling.xaxis.set_major_formatter(DateFormatter('%Y'))
        self.ax_rolling.legend(loc='upper right')
        # Tracé de la distribution (histogramme)
        from matplotlib.ticker import PercentFormatter
        n, bins, patches = self.ax_rolling_hist.hist(rolling_return.dropna() * 100, bins=30, color='#4F8EF7', alpha=0.75, density=True)
        # Réduit la largeur des barres à 90% pour laisser un espace
        for patch in patches:
            current_width = patch.get_width()
            patch.set_width(current_width * 0.9)
            patch.set_x(patch.get_x() + current_width * 0.05)
        self.ax_rolling_hist.set_title('Distribution des rendements glissants')
        self.ax_rolling_hist.set_xlabel('Rendement glissant (%)')
        self.ax_rolling_hist.set_ylabel('Densité (%)')
        self.ax_rolling_hist.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))
        self.ax_rolling_hist.grid(True, linestyle='--')
        # Ajout du label de probabilité inférieur à 0% dans la légende
        rolling_return_clean = rolling_return.dropna()
        prob_neg = (rolling_return_clean < 0).mean()
        label_text = f"P(R < 0%) = {prob_neg:.1%}"
        # Ligne verticale à x=0 pour support du label
        self.ax_rolling_hist.axvline(0, color='red', linestyle='--', linewidth=2, label=label_text)
        self.ax_rolling_hist.legend(loc='upper right')
        self.figure_rolling.tight_layout()
        self.canvas_rolling.draw()


    def plot_drawdown(self, df, freq):
        """
        Affiche le drawdown dans l'axe inférieur de l'onglet 'cours'.
        """
        import numpy as np
        ax = self.ax_drawdown
        ax.clear()
        if 'Close' not in df.columns or df['Close'].isnull().all():
            ax.set_title('Drawdown')
            ax.set_ylabel('Drawdown')
            ax.grid(True)
            return
        prices = df['Close']
        # Calcul du drawdown
        cummax = prices.cummax()
        drawdown = (prices / cummax) - 1
        ax.plot(df['Date'], drawdown, color='crimson', label='Drawdown')
        ax.fill_between(df['Date'], drawdown, 0, where=(drawdown<0), color='crimson', alpha=0.3)
        ax.set_title('Drawdown')
        ax.set_ylabel('Drawdown')
        ax.grid(True)
        ax.set_ylim(-1, 0)
        ax.legend(loc='lower left')

    def update_graphs(self):
        from scipy import stats
        actif = self.asset_select.currentText()
        file_path = os.path.join(os.path.dirname(__file__), actif + '.xlsx')
        title_price = f'Cours de bourse {actif}'
        title_return = f'Rendements {actif}'
        if not os.path.exists(file_path):
            return
        df = pd.read_excel(file_path, engine='openpyxl', decimal=',')
        df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y')
        df = df.sort_values('Date')
        # Filtrage selon les dates si "Toute la série" n'est pas cochée
        if not self.all_series_checkbox.isChecked():
            date_debut = self.date_start.date().toPyDate()
            date_fin = self.date_end.date().toPyDate()
            df = df[(df['Date'] >= pd.Timestamp(date_debut)) & (df['Date'] <= pd.Timestamp(date_fin))]
        # Sinon, on garde toute la série

        # Rééchantillonnage selon la fréquence sélectionnée
        freq_map = {
            'Journalier': None,
            'Hebdomadaire': 'W',
            'Mensuelle': 'ME',
            'Annuelle': 'YE',
        }
        freq = self.freq_select.currentText()
        if freq_map[freq]:
            df = df.set_index('Date').resample(freq_map[freq]).last().reset_index()

        # Graphe des prix
        self.ax_price.clear()
        self.ax_price.plot(df['Date'], df['Close'], linestyle='-')
        self.ax_price.set_title(title_price + f" ({freq})")
        self.ax_price.set_ylabel('Close')
        self.ax_price.grid(True)
        # Drawdown
        self.plot_drawdown(df, freq)
        self.ax_drawdown.set_xlabel('Date')
        self.figure.tight_layout()
        self.canvas.draw()

        # Calcul du rendement selon le choix
        import numpy as np
        return_type = self.return_type_select.currentText()
        if return_type == "Simple":
            returns = df['Close'].pct_change()
        else:
            returns = np.log(df['Close'] / df['Close'].shift(1))

        # Graphe des rendements (série temporelle)
        self.ax_return.clear()
        import seaborn as sns
        sns.lineplot(x=df['Date'], y=returns, ax=self.ax_return, color='tab:green', linestyle='-', label='Rendement')

        # Calcul des +/- 2 écarts-types
        std = returns.std(skipna=True)
        self.ax_return.axhline(2*std, color='red', linestyle='--', linewidth=1, label='+2 écarts-types')
        self.ax_return.axhline(-2*std, color='red', linestyle='--', linewidth=1, label='-2 écarts-types')

        self.ax_return.set_title(title_return + f" ({freq})")
        self.ax_return.set_xlabel('Date')
        self.ax_return.set_ylabel('Rendement ' + freq.lower())
        self.ax_return.grid(True)
        self.ax_return.legend()

        # Histogramme des rendements (en dessous) avec seaborn et ajustement Student-t
        import seaborn as sns
        returns_clean = returns.dropna()  # RESTE EN DÉCIMAL
        self.ax_hist.clear()
        import scipy.stats as stats
        sns.histplot(returns_clean * 100, bins=50, kde=True, ax=self.ax_hist, color='tab:orange', stat="density", label="Kernel Density Estimate", edgecolor=None, alpha=0.5, shrink=0.95)
        df_fit, loc, scale = stats.t.fit(returns_clean * 100)
        x = pd.Series(returns_clean * 100).sort_values()
        y = stats.t.pdf(x, df_fit, loc, scale)
        self.ax_hist.plot(x, y, color='blue', linestyle="--", label="Ajustement Student-t")
        self.ax_hist.legend()
        self.ax_hist.set_title("Histogramme des rendements (%)")
        self.ax_hist.grid(True)
        self.ax_hist.set_xlabel("Rendement (%)")
        # Augmenter la fréquence des ticks de l'axe x
        import numpy as np
        x_min, x_max = self.ax_hist.get_xlim()
        step = 2  # pas de 2%
        ticks = np.arange(np.floor(x_min), np.ceil(x_max)+step, step)
        self.ax_hist.set_xticks(ticks)
        self.ax_hist.tick_params(axis='x')  # Optionnel : incliner les labels pour la lisibilité
        self.ax_hist.set_ylabel("Densité")
        param_annotation = f"Student-t params\ndf = {df_fit:.2f}\nμ = {loc:.4f}\nσ = {scale:.4f}"
        self.ax_hist.text(0.98, 0.05, param_annotation, transform=self.ax_hist.transAxes,
                          fontsize=10, verticalalignment='bottom', horizontalalignment='right',
                          bbox=dict(facecolor='white', edgecolor='black'))
        self.ax_return.grid(True)
        self.figure_return.tight_layout()
        self.canvas_return.draw()
        self.canvas_return.flush_events()
        self.canvas_return.updateGeometry()
        if self.tab_return.layout() is not None:
            self.tab_return.layout().activate()
        self.tab_return.update()


        # --- Nuage de points (rendement t vs t-1) ---
        self.ax_scatter.clear()
        # On utilise returns_clean pour avoir les rendements filtrés et propres
        returns_t = returns_clean[1:]
        returns_tm1 = returns_clean[:-1]
        self.ax_scatter.scatter(returns_tm1, returns_t, alpha=0.5, color='tab:blue')
        self.ax_scatter.margins(0.1)  # Pour éviter que les points soient collés aux bords
        # Calcul manuel des coefficients de la droite de régression
        x = returns_tm1.values
        y = returns_t.values
        x_mean = x.mean()
        y_mean = y.mean()
        a = ((x - x_mean) * (y - y_mean)).sum() / ((x - x_mean)**2).sum()
        b = y_mean - a * x_mean
        # Tracé de la droite
        x_line = np.array([x.min(), x.max()])
        y_line = a * x_line + b
        self.ax_scatter.plot(x_line, y_line, color='red', lw=2, label=f"Droite de régression : y = {a:.2f}x + {b:.2g}")
        self.ax_scatter.set_xlabel('Rendement t-1')
        self.ax_scatter.set_ylabel('Rendement t')
        self.ax_scatter.set_title(f'Nuage de points : Rendement t vs t-1 ({freq})')
        self.ax_scatter.legend()
        self.ax_scatter.grid(True)
        self.figure_scatter.tight_layout()
        self.canvas_scatter.draw()
        self.canvas_scatter.flush_events()
        self.canvas_scatter.flush_events()

        # --- QQ-Plot ---
        self.ax_qq.clear()
        stats.probplot(returns_clean, dist="norm", plot=self.ax_qq)
        for line in self.ax_qq.get_lines():
            if line.get_linestyle() == 'None':  # Ce sont les points
                line.set_color('#1f77b4')  # Bleu clair
        self.ax_qq.grid(True)  # Affiche la grille sur le QQ-plot
        self.figure_qq.tight_layout()
        self.canvas_qq.draw()
        self.canvas_qq.flush_events()
        # Test de Jarque-Bera
        from scipy.stats import jarque_bera
        jb_stat, jb_pvalue = jarque_bera(returns_clean)
        normality = "Normal" if jb_pvalue > 0.05 else "Non-Normal"
        label_jb = f"JB = {jb_stat:.2f}\n{normality}"
        # Affichage en haut à droite du QQ-Plot
        self.ax_qq.text(0.98, 0.02, label_jb, transform=self.ax_qq.transAxes,
                        fontsize=12, color='black', ha='right', va='bottom',
                        bbox=dict(facecolor='white', alpha=0.7, edgecolor='gray'))
        self.ax_qq.set_title('QQ-Plot des rendements')

        # ACF (à gauche du bas)
        self.ax_acf.clear()
        plot_acf(returns_clean, ax=self.ax_acf, lags=10, zero=False, alpha=0.05)
        self.ax_acf.set_title('ACF des rendements')

        # PACF (à droite du bas)
        self.ax_pacf.clear()
        plot_pacf(returns_clean, ax=self.ax_pacf, lags=10, zero=False, alpha=0.05, method='ywm')
        self.ax_pacf.set_title('PACF des rendements')

        self.figure_qq.tight_layout()
        self.canvas_qq.draw()
        self.canvas_qq.flush_events()
        self.canvas_qq.updateGeometry()
        if self.tab_qq.layout() is not None:
            self.tab_qq.layout().activate()
        self.tab_qq.update()

        # --- Statistiques ---
        import numpy as np
        from scipy.stats import skew, kurtosis
        mean = returns_clean.mean()
        median = returns_clean.median()
        minimum = returns_clean.min()
        maximum = returns_clean.max()
        std = returns_clean.std()
        skewness = skew(returns_clean)
        kurt = kurtosis(returns_clean)
        var_5 = returns_clean.quantile(0.05)
        cvar_5 = returns_clean[returns_clean <= var_5].mean()
        # Maximum Drawdown
        cumulative = (1 + returns_clean).cumprod()
        running_max = cumulative.cummax()
        drawdown = (cumulative - running_max) / running_max
        max_drawdown = drawdown.min()
        plus_3std = 3 * std
        minus_3std = -3 * std
        # Remplissage du tableau
        stats_labels = [
            "Moyenne",
            "Médiane",
            "Minimum",
            "Maximum",
            "Écart-type",
            "Skewness",
            "Kurtosis",
            "VaR 5%",
            "CVaR 5%",
            "Maximum Drawdown",
        ]
        stats_values = [
            mean,
            median,
            minimum,
            maximum,
            std,
            skewness,
            kurt,
            var_5,
            cvar_5,
            max_drawdown,
        ]
        from PyQt5.QtGui import QFont
        font_stats = QFont()
        font_stats.setPointSize(15)  # Taille de police augmentée
        for i, (label, value) in enumerate(zip(stats_labels, stats_values)):
            item_label = QTableWidgetItem(label)
            item_label.setFont(font_stats)
            self.stats_table.setItem(i, 0, item_label)
            if label in ["Skewness", "Kurtosis"]:
                item_value = QTableWidgetItem(f"{value:.4f}")
            else:
                percent_value = value * 100
                item_value = QTableWidgetItem(f"{percent_value:.4f} %")
            item_value.setFont(font_stats)
            self.stats_table.setItem(i, 1, item_value)
        # Ajout ±3 écart-types en pourcentage sur la ligne Écart-type
        std_percent = std * 100
        minus_3std_percent = minus_3std * 100
        plus_3std_percent = plus_3std * 100
        item_std = QTableWidgetItem(f"{std_percent:.4f} % (±3σ: [{minus_3std_percent:.4f} %, {plus_3std_percent:.4f} %])")
        item_std.setFont(font_stats)
        self.stats_table.setItem(4, 1, item_std)


    def plot_rolling_moments(self):
        import numpy as np
        import pandas as pd
        from scipy.stats import skew, kurtosis
        from matplotlib.dates import DateFormatter
        actif = self.asset_select.currentText()
        file_path = os.path.join(os.path.dirname(__file__), actif + '.xlsx')
        if not os.path.exists(file_path):
            return
        df = pd.read_excel(file_path, engine='openpyxl', decimal=',')
        df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y')
        df = df.sort_values('Date')
        # Filtrage selon la période choisie
        if not self.all_series_checkbox.isChecked():
            date_debut = self.date_start.date().toPyDate()
            date_fin = self.date_end.date().toPyDate()
            df = df[(df['Date'] >= pd.Timestamp(date_debut)) & (df['Date'] <= pd.Timestamp(date_fin))]
        if 'Close' in df.columns:
            prices = df['Close']
        else:
            prices = df.iloc[:, 1]  # 2e colonne si pas de nom
        returns = prices.pct_change().dropna()
        dates = df['Date'].iloc[1:]
        window = self.rolling_window_spin.value()
        # Moments glissants
        rolling_mean = returns.rolling(window).mean()
        rolling_vol = returns.rolling(window).std()
        rolling_skew = returns.rolling(window).apply(skew, raw=True)
        rolling_kurt = returns.rolling(window).apply(kurtosis, raw=True)
        # Affichage
        self.ax_mean.clear()
        self.ax_vol.clear()
        self.ax_skew.clear()
        self.ax_kurt.clear()
        self.ax_mean.plot(dates, rolling_mean, color='black', label=f"Mean (rolling window {window} days)")
        self.ax_mean.axhline(0, color='gray', lw=0.7)
        self.ax_mean.set_title('Time-varying mean')
        self.ax_mean.set_ylabel('Mean')
        self.ax_mean.legend(loc='upper right')
        self.ax_vol.plot(dates, rolling_vol, color='black', label=f"Volatility (rolling window {window} days)")
        self.ax_vol.set_title('Time-varying volatility')
        self.ax_vol.set_ylabel('Volatility')
        self.ax_vol.legend(loc='upper right')
        self.ax_skew.plot(dates, rolling_skew, color='black', label=f"Skewness (rolling window {window} days)")
        self.ax_skew.axhline(0, color='gray', lw=0.7)
        self.ax_skew.set_title('Time-varying skewness')
        self.ax_skew.set_ylabel('Skewness')
        self.ax_skew.legend(loc='upper right')
        self.ax_kurt.plot(dates, rolling_kurt, color='black', label=f"Kurtosis (rolling window {window} days)")
        self.ax_kurt.set_title('Time-varying kurtosis')
        self.ax_kurt.set_ylabel('Kurtosis')
        self.ax_kurt.legend(loc='upper right')
        for ax in [self.ax_mean, self.ax_vol, self.ax_skew, self.ax_kurt]:
            ax.set_xlabel('Date')
            ax.grid(True)
            ax.xaxis.set_major_formatter(DateFormatter('%Y'))
        self.figure_moments.tight_layout()
        self.canvas_moments.draw()

# --- Fonctions utilitaires pour la sélection cumulative return ---
from PyQt5.QtCore import Qt

def cumret_select_all(self):
    for i in range(self.cumret_asset_list.count()):
        self.cumret_asset_list.item(i).setCheckState(Qt.Checked)
    self.update_cumret_selected_label()

def cumret_deselect_all(self):
    for i in range(self.cumret_asset_list.count()):
        self.cumret_asset_list.item(i).setCheckState(Qt.Unchecked)
    self.update_cumret_selected_label()

def update_cumret_selected_label(self):
    # Ne rien afficher
    self.cumret_selected_label.setText("")

# Ajout des méthodes à MainWindow
MainWindow.cumret_select_all = cumret_select_all
MainWindow.cumret_deselect_all = cumret_deselect_all
MainWindow.update_cumret_selected_label = update_cumret_selected_label

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec_())

plt.show()