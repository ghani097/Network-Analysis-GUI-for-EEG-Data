"""
Network Level Analysis - PyQt5 GUI
===================================

A modern graphical interface for PLI network analysis.

Usage:
    python network_analysis_gui.py
"""

import os
import sys
import threading
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QPushButton, QLineEdit, QCheckBox, QComboBox,
    QTableWidget, QTableWidgetItem, QTextEdit, QProgressBar,
    QFileDialog, QMessageBox, QSplitter, QFrame, QHeaderView,
    QGridLayout, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette

import pandas as pd

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from network_analysis import NetworkAnalysis


class AnalysisWorker(QThread):
    """Background worker thread for analysis."""
    progress = pyqtSignal(str, int)  # message, percentage
    finished = pyqtSignal(bool, str)  # success, message

    def __init__(self, config):
        super().__init__()
        self.config = config

    def run(self):
        try:
            def callback(msg, pct):
                self.progress.emit(msg, pct if pct is not None else -1)

            analysis = NetworkAnalysis(
                input_file=self.config['input_file'],
                output_dir=self.config['output_dir'],
                adjust_baseline=self.config['adjust_baseline'],
                groups=self.config['groups'],
                sessions=self.config['sessions'],
                networks=self.config['networks'],
                frequency_bands=self.config['frequency_bands'],
                callback=callback
            )

            analysis.run()
            self.finished.emit(True, f"Analysis complete!\nResults saved to: {self.config['output_dir']}")

        except Exception as e:
            self.finished.emit(False, str(e))


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLI Network Analysis")
        self.setMinimumSize(900, 700)
        self.resize(1000, 800)

        # Data
        self.df = None
        self.worker = None

        # Build UI
        self._build_ui()

        # Load default file if exists
        default_file = Path(__file__).parent / "PLI-UK-Both-Groups-UP3.xlsx"
        if default_file.exists():
            self.input_edit.setText(str(default_file))
            self._load_file(str(default_file))

    def _build_ui(self):
        """Build the user interface."""
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)

        # === Title ===
        title = QLabel("PLI Network Analysis")
        title.setFont(QFont("Segoe UI", 16, QFont.Bold))
        layout.addWidget(title)

        subtitle = QLabel("Analyze Phase Lag Index across brain networks")
        subtitle.setStyleSheet("color: gray;")
        layout.addWidget(subtitle)

        # === File Selection ===
        file_group = QGroupBox("Input File")
        file_layout = QHBoxLayout(file_group)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Select Excel file...")
        file_layout.addWidget(self.input_edit, 1)

        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_file)
        file_layout.addWidget(browse_btn)

        load_btn = QPushButton("Load")
        load_btn.clicked.connect(lambda: self._load_file(self.input_edit.text()))
        load_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        file_layout.addWidget(load_btn)

        layout.addWidget(file_group)

        # === Splitter for data preview and options ===
        splitter = QSplitter(Qt.Vertical)

        # --- Data Preview Section ---
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.setContentsMargins(0, 0, 0, 0)

        # Data info label
        self.data_info = QLabel("No file loaded")
        self.data_info.setStyleSheet("font-weight: bold; color: #333;")
        preview_layout.addWidget(self.data_info)

        # Data structure summary
        struct_layout = QHBoxLayout()

        self.groups_label = QLabel("Groups: -")
        self.groups_label.setStyleSheet("background: #e3f2fd; padding: 5px; border-radius: 3px;")
        struct_layout.addWidget(self.groups_label)

        self.sessions_label = QLabel("Sessions: -")
        self.sessions_label.setStyleSheet("background: #e8f5e9; padding: 5px; border-radius: 3px;")
        struct_layout.addWidget(self.sessions_label)

        self.networks_label = QLabel("Networks: -")
        self.networks_label.setStyleSheet("background: #fff3e0; padding: 5px; border-radius: 3px;")
        struct_layout.addWidget(self.networks_label)

        self.freq_label = QLabel("Frequencies: -")
        self.freq_label.setStyleSheet("background: #fce4ec; padding: 5px; border-radius: 3px;")
        struct_layout.addWidget(self.freq_label)

        preview_layout.addLayout(struct_layout)

        # Data preview table
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setMaximumHeight(150)
        preview_layout.addWidget(self.preview_table)

        splitter.addWidget(preview_widget)

        # --- Options Section ---
        options_widget = QWidget()
        options_layout = QVBoxLayout(options_widget)
        options_layout.setContentsMargins(0, 0, 0, 0)

        # Analysis options row
        opt_group = QGroupBox("Analysis Options")
        opt_layout = QHBoxLayout(opt_group)

        # Baseline checkbox
        self.baseline_check = QCheckBox("Adjust for Baseline (ANCOVA)")
        self.baseline_check.setChecked(True)
        self.baseline_check.stateChanged.connect(self._on_baseline_change)
        opt_layout.addWidget(self.baseline_check)

        self.baseline_info = QLabel("Pre session used as covariate")
        self.baseline_info.setStyleSheet("color: green; font-style: italic;")
        opt_layout.addWidget(self.baseline_info)

        opt_layout.addStretch()

        # Output directory
        opt_layout.addWidget(QLabel("Output:"))
        self.output_edit = QLineEdit("analysis_output")
        self.output_edit.setMaximumWidth(200)
        opt_layout.addWidget(self.output_edit)

        options_layout.addWidget(opt_group)

        # Selection checkboxes
        sel_layout = QHBoxLayout()

        # Groups
        self.groups_group = QGroupBox("Groups")
        groups_layout = QVBoxLayout(self.groups_group)
        self.group_checks = {}
        sel_layout.addWidget(self.groups_group)

        # Sessions
        self.sessions_group = QGroupBox("Sessions")
        sessions_layout = QVBoxLayout(self.sessions_group)
        self.session_checks = {}
        sel_layout.addWidget(self.sessions_group)

        # Networks
        self.networks_group = QGroupBox("Networks")
        networks_layout = QVBoxLayout(self.networks_group)
        self.network_checks = {}
        sel_layout.addWidget(self.networks_group)

        # Frequency Bands
        self.freq_group = QGroupBox("Frequency Bands")
        freq_layout = QVBoxLayout(self.freq_group)
        self.freq_checks = {}
        sel_layout.addWidget(self.freq_group)

        options_layout.addLayout(sel_layout)

        splitter.addWidget(options_widget)

        # --- Log Section ---
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 0, 0, 0)

        log_header = QHBoxLayout()
        log_header.addWidget(QLabel("Output Log"))

        clear_btn = QPushButton("Clear")
        clear_btn.setMaximumWidth(60)
        clear_btn.clicked.connect(lambda: self.log_text.clear())
        log_header.addWidget(clear_btn)

        log_layout.addLayout(log_header)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        self.log_text.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4;")
        log_layout.addWidget(self.log_text)

        splitter.addWidget(log_widget)

        # Set splitter sizes
        splitter.setSizes([180, 180, 300])
        layout.addWidget(splitter, 1)

        # === Progress Bar ===
        progress_layout = QHBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Ready")
        self.status_label.setMinimumWidth(150)
        progress_layout.addWidget(self.status_label)

        layout.addLayout(progress_layout)

        # === Buttons ===
        btn_layout = QHBoxLayout()

        self.run_btn = QPushButton("Run Analysis")
        self.run_btn.setFont(QFont("Segoe UI", 10, QFont.Bold))
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 10px 30px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #ccc;
            }
        """)
        self.run_btn.clicked.connect(self._run_analysis)
        btn_layout.addWidget(self.run_btn)

        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 10px 20px;
                border-radius: 5px;
            }
            QPushButton:disabled {
                background-color: #ccc;
            }
        """)
        btn_layout.addWidget(self.stop_btn)

        btn_layout.addStretch()

        open_btn = QPushButton("Open Output Folder")
        open_btn.clicked.connect(self._open_output)
        btn_layout.addWidget(open_btn)

        layout.addLayout(btn_layout)

        # Initialize with defaults
        self._init_checkboxes()

    def _init_checkboxes(self):
        """Initialize checkbox groups - all populated dynamically from loaded Excel file."""
        # Groups - will be populated dynamically from loaded Excel file
        self.groups_placeholder = QLabel("Load file to see groups")
        self.groups_placeholder.setStyleSheet("color: gray; font-style: italic;")
        self.groups_group.layout().addWidget(self.groups_placeholder)

        # Sessions - will be populated dynamically from loaded Excel file
        self.sessions_placeholder = QLabel("Load file to see sessions")
        self.sessions_placeholder.setStyleSheet("color: gray; font-style: italic;")
        self.sessions_group.layout().addWidget(self.sessions_placeholder)

        # Networks - will be populated dynamically from loaded Excel file
        self.networks_placeholder = QLabel("Load file to see networks")
        self.networks_placeholder.setStyleSheet("color: gray; font-style: italic;")
        self.networks_group.layout().addWidget(self.networks_placeholder)

        # Frequencies - will be populated dynamically from loaded Excel file
        self.freq_placeholder = QLabel("Load file to see frequencies")
        self.freq_placeholder.setStyleSheet("color: gray; font-style: italic;")
        self.freq_group.layout().addWidget(self.freq_placeholder)

    def _browse_file(self):
        """Browse for input file."""
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "",
            "Excel Files (*.xlsx *.xls *.xlsm);;All Files (*)"
        )
        if file:
            self.input_edit.setText(file)
            self._load_file(file)

    def _load_file(self, filepath):
        """Load and preview the Excel file."""
        if not filepath or not Path(filepath).exists():
            return

        try:
            self.df = pd.read_excel(filepath)

            # Ensure categorical columns are strings for consistent handling
            for col in ['Group', 'Session', 'Network', 'FrequencyTag']:
                if col in self.df.columns:
                    self.df[col] = self.df[col].astype(str)

            # Update data info
            self.data_info.setText(f"Loaded: {Path(filepath).name} ({len(self.df)} rows, {len(self.df.columns)} columns)")
            self.data_info.setStyleSheet("font-weight: bold; color: green;")

            # Update structure labels
            if 'Group' in self.df.columns:
                groups = list(self.df['Group'].unique())
                self.groups_label.setText(f"Groups: {', '.join(groups)}")
                self._create_group_checkboxes(groups)

            if 'Session' in self.df.columns:
                sessions = list(self.df['Session'].unique())
                self.sessions_label.setText(f"Sessions: {', '.join(sessions)}")
                self._create_session_checkboxes(sessions)

            if 'Network' in self.df.columns:
                networks = list(self.df['Network'].unique())
                self.networks_label.setText(f"Networks: {', '.join(networks)}")
                self._create_network_checkboxes(networks)

            if 'FrequencyTag' in self.df.columns:
                freqs = list(self.df['FrequencyTag'].unique())
                self.freq_label.setText(f"Frequencies: {', '.join(freqs)}")
                self._create_freq_checkboxes(freqs)

            # Update preview table
            self._update_preview_table()

            self._log(f"File loaded: {filepath}")
            self._log(f"  Shape: {self.df.shape}")
            self._log(f"  Columns: {', '.join(self.df.columns)}")

        except Exception as e:
            self.data_info.setText(f"Error loading file: {e}")
            self.data_info.setStyleSheet("font-weight: bold; color: red;")
            QMessageBox.critical(self, "Error", f"Failed to load file:\n{e}")

    def _update_checkboxes(self, checks_dict, values, default=None):
        """Update checkboxes to match available values."""
        for name, cb in checks_dict.items():
            if name in values:
                cb.setEnabled(True)
                if default:
                    cb.setChecked(name in default)
                else:
                    cb.setChecked(True)
            else:
                cb.setEnabled(False)
                cb.setChecked(False)

    def _create_group_checkboxes(self, groups):
        """Dynamically create group checkboxes from loaded data."""
        # Remove placeholder if exists
        if hasattr(self, 'groups_placeholder') and self.groups_placeholder is not None:
            self.groups_placeholder.deleteLater()
            self.groups_placeholder = None

        # Remove existing group checkboxes
        for cb in list(self.group_checks.values()):
            cb.deleteLater()
        self.group_checks.clear()

        # Create new checkboxes for each group from the Excel file
        for group in groups:
            cb = QCheckBox(str(group))
            cb.setChecked(True)
            self.groups_group.layout().addWidget(cb)
            self.group_checks[str(group)] = cb

    def _create_session_checkboxes(self, sessions):
        """Dynamically create session checkboxes from loaded data."""
        # Remove placeholder if exists
        if hasattr(self, 'sessions_placeholder') and self.sessions_placeholder is not None:
            self.sessions_placeholder.deleteLater()
            self.sessions_placeholder = None

        # Remove existing session checkboxes
        for cb in list(self.session_checks.values()):
            cb.deleteLater()
        self.session_checks.clear()

        # Create new checkboxes for each session from the Excel file
        for session in sessions:
            cb = QCheckBox(str(session))
            cb.setChecked(True)
            self.sessions_group.layout().addWidget(cb)
            self.session_checks[str(session)] = cb

    def _create_network_checkboxes(self, networks):
        """Dynamically create network checkboxes from loaded data."""
        # Remove placeholder if exists
        if hasattr(self, 'networks_placeholder') and self.networks_placeholder is not None:
            self.networks_placeholder.deleteLater()
            self.networks_placeholder = None

        # Remove existing network checkboxes
        for cb in list(self.network_checks.values()):
            cb.deleteLater()
        self.network_checks.clear()

        # Create new checkboxes for each network from the Excel file
        for network in networks:
            cb = QCheckBox(str(network))
            cb.setChecked(True)
            self.networks_group.layout().addWidget(cb)
            self.network_checks[str(network)] = cb

    def _create_freq_checkboxes(self, freqs):
        """Dynamically create frequency band checkboxes from loaded data."""
        # Remove placeholder if exists
        if hasattr(self, 'freq_placeholder') and self.freq_placeholder is not None:
            self.freq_placeholder.deleteLater()
            self.freq_placeholder = None

        # Remove existing frequency checkboxes
        for cb in list(self.freq_checks.values()):
            cb.deleteLater()
        self.freq_checks.clear()

        # Create new checkboxes for each frequency band from the Excel file
        for freq in freqs:
            cb = QCheckBox(str(freq))
            cb.setChecked(True)
            self.freq_group.layout().addWidget(cb)
            self.freq_checks[str(freq)] = cb

    def _update_preview_table(self):
        """Update the data preview table."""
        if self.df is None:
            return

        # Show first 10 rows
        preview = self.df.head(10)

        self.preview_table.setRowCount(len(preview))
        self.preview_table.setColumnCount(len(preview.columns))
        self.preview_table.setHorizontalHeaderLabels(preview.columns.tolist())

        for i, row in preview.iterrows():
            for j, val in enumerate(row):
                item = QTableWidgetItem(str(val) if not pd.isna(val) else "")
                self.preview_table.setItem(i, j, item)

        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def _on_baseline_change(self):
        """Handle baseline checkbox change."""
        if self.baseline_check.isChecked():
            self.baseline_info.setText("Pre session used as covariate")
            self.baseline_info.setStyleSheet("color: green; font-style: italic;")
        else:
            self.baseline_info.setText("All sessions included in analysis")
            self.baseline_info.setStyleSheet("color: blue; font-style: italic;")

    def _get_selected(self, checks_dict):
        """Get list of selected checkbox values."""
        return [k for k, cb in checks_dict.items() if cb.isChecked() and cb.isEnabled()]

    def _validate(self):
        """Validate inputs before running."""
        if not self.input_edit.text():
            QMessageBox.warning(self, "Warning", "Please select an input file.")
            return False

        if not Path(self.input_edit.text()).exists():
            QMessageBox.warning(self, "Warning", "Input file not found.")
            return False

        if not self._get_selected(self.group_checks):
            QMessageBox.warning(self, "Warning", "Select at least one group.")
            return False

        if not self._get_selected(self.network_checks):
            QMessageBox.warning(self, "Warning", "Select at least one network.")
            return False

        if not self._get_selected(self.freq_checks):
            QMessageBox.warning(self, "Warning", "Select at least one frequency band.")
            return False

        return True

    def _log(self, msg):
        """Add message to log."""
        self.log_text.append(msg)
        # Scroll to bottom
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _run_analysis(self):
        """Start the analysis."""
        if not self._validate():
            return

        # Disable UI
        self.run_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Running...")
        self.status_label.setStyleSheet("color: blue; font-weight: bold;")

        # Clear log
        self.log_text.clear()

        # Prepare config
        config = {
            'input_file': self.input_edit.text(),
            'output_dir': self.output_edit.text(),
            'adjust_baseline': self.baseline_check.isChecked(),
            'groups': self._get_selected(self.group_checks),
            'sessions': self._get_selected(self.session_checks),
            'networks': self._get_selected(self.network_checks),
            'frequency_bands': self._get_selected(self.freq_checks)
        }

        self._log("=" * 50)
        self._log("Starting Analysis")
        self._log("=" * 50)
        self._log(f"Input: {config['input_file']}")
        self._log(f"Output: {config['output_dir']}")
        self._log(f"Baseline: {'Yes' if config['adjust_baseline'] else 'No'}")
        self._log(f"Groups: {', '.join(config['groups'])}")
        self._log(f"Networks: {', '.join(config['networks'])}")
        self._log(f"Frequencies: {', '.join(config['frequency_bands'])}")
        self._log("=" * 50)

        # Start worker thread
        self.worker = AnalysisWorker(config)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.start()

    def _on_progress(self, msg, pct):
        """Handle progress updates from worker."""
        self._log(msg)
        if pct >= 0:
            self.progress_bar.setValue(pct)
            self.status_label.setText(f"Running... {pct}%")

    def _on_finished(self, success, message):
        """Handle analysis completion."""
        self.run_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        if success:
            self.progress_bar.setValue(100)
            self.status_label.setText("Complete!")
            self.status_label.setStyleSheet("color: green; font-weight: bold;")
            QMessageBox.information(self, "Success", message)
        else:
            self.status_label.setText("Error!")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self._log(f"\nERROR: {message}")
            QMessageBox.critical(self, "Error", message)

    def _open_output(self):
        """Open the output folder."""
        path = Path(self.output_edit.text())
        if path.exists():
            if sys.platform == 'win32':
                os.startfile(str(path))
            elif sys.platform == 'darwin':
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        else:
            QMessageBox.information(self, "Info", f"Output folder doesn't exist yet:\n{path}")


def main():
    app = QApplication(sys.argv)

    # Set application style
    app.setStyle('Fusion')

    # Optional: Dark palette
    # palette = QPalette()
    # palette.setColor(QPalette.Window, QColor(53, 53, 53))
    # app.setPalette(palette)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
