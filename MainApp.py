import numpy as np
import os
import sys
import xlsxwriter as xl
from PyQt5.QtCore import Qt, QRect, QObject, QThreadPool, QRunnable, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QMainWindow,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QGroupBox,
    QCheckBox,
    QSizePolicy,
    QComboBox,
    QRadioButton,
    QMessageBox,
    QStackedWidget,
    QButtonGroup,
    QProgressBar
)
from PyQt5.QtGui import QIcon
import pandas as pd
from pathlib import Path

basedir = os.path.dirname(__file__)


class WorkerSignals(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    error = pyqtSignal()


class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        self.kwargs['progress_callback'] = self.signals.progress

    @pyqtSlot()
    def run(self):
        try:
            self.fn(*self.args, **self.kwargs)
        except:
            self.signals.error.emit()
        else:
            self.signals.finished.emit()


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.threadpool = QThreadPool()
        self.defaultDataset = 'Achievement'
        self.currentDataset = self.defaultDataset
        self.currentYears = []
        self.currentLevel = None
        self.setGeometry(50, 100, 1700, 700)
        self.setFixedSize(self.size())
        self.setWindowTitle('Unity Insights - QOF')
        self.setWindowIcon(QIcon(os.path.join(basedir, 'Dependencies/logoU.png')))

        with open(os.path.join(basedir, 'Dependencies/style.qss')) as f:
            style = f.read()
            self.setStyleSheet(style)

        self.prevFrame = self.ReadCsv(os.path.join(basedir, 'Dependencies/AppData/AppPrevalenceFrame.csv'),
                                      index=[0, 1], header=[0, 1])
        self.achFrame = self.ReadCsv(os.path.join(basedir, 'Dependencies/AppData/AppAchievementFrame.csv'),
                                     index=[0, 1], header=[0, 1])
        self.mapFrame = self.ReadCsv(os.path.join(basedir, 'Dependencies/AppData/AppMapFrame.csv'), index=0, header=0)

        self.AchievementDefinitions = self.ReadCsv(os.path.join(basedir, 'Dependencies/AppData/AchievementDefinitions.csv'))

        self.PrevalenceDefinitions = self.ReadCsv(os.path.join(basedir, 'Dependencies/AppData/PrevalenceDefinitions.csv'))

        self.ui_init()  # Calls UI Function

    # Contains all UI design & layout
    def ui_init(self):

        # Main Grid
        self.grid = QGridLayout()
        self.grid.setSpacing(10)
        self.grid.setContentsMargins(10, 10, 10, 10)

        widget = QWidget()
        widget.setLayout(self.grid)
        self.setCentralWidget(widget)

        # Left hand 'select years' Group Box
        self.year_group = QGroupBox('Select years...')
        self.grid.addWidget(self.year_group, 0, 0, 1, 1)
        self.year_group.setObjectName('MainBox')

        # Vertical Layout for whole group
        self.yearVerticalLayout = QVBoxLayout()
        self.year_group.setLayout(self.yearVerticalLayout)

        # Widget to parent the Vertical Box to
        self.verticalDateWidget = QWidget(self.year_group)
        self.yearVerticalLayout.addWidget(self.verticalDateWidget)
        # Vertical Box for layout, Adding date checkboxes
        self.date_vbox = QVBoxLayout(self.verticalDateWidget)
        self.date_vbox.addSpacing(20)
        dates = ['2021-22', '2020-21', '2019-20', '2018-19', '2017-18', '2016-17']
        for date in dates:
            cb_dates = QCheckBox(date)
            self.date_vbox.addWidget(cb_dates)
        self.selectAll = QCheckBox('Select all')
        self.selectAll.setObjectName('selectAll')
        self.selectAll.clicked.connect(self.SelectAllYears)
        self.date_vbox.addWidget(self.selectAll)

        # Confirm button below date checkboxes
        self.date_confirm = QPushButton('Confirm')
        self.yearVerticalLayout.addWidget(self.date_confirm)
        self.date_confirm.setMinimumSize(10, 40)
        self.date_confirm.clicked.connect(self.YearsOnClick)

        # Text box to show which dates have been selected
        self.date_list = QLabel('Years Selected:')
        self.yearVerticalLayout.addWidget(self.date_list)
        self.date_list.setAlignment(Qt.AlignTop)
        self.date_list.setStyleSheet('Padding:10px')

        # Select Place Group
        self.placeSelectGroup = QGroupBox('Select site(s)...')
        self.grid.addWidget(self.placeSelectGroup, 0, 1, 1, 2)
        self.placeSelectGroup.setGeometry(0, 0, 640, 680)
        self.placeSelectGroup.setObjectName('MainBox')

        self.placeVerticalLayout = QVBoxLayout()
        self.placeSelectGroup.setLayout(self.placeVerticalLayout)
        self.placeVerticalLayout.addSpacing(25)

        # Widget to parent horizontal layout box to
        self.horizontalLevelWidget = QWidget()
        self.placeVerticalLayout.addWidget(self.horizontalLevelWidget)
        self.horizontalLevelWidget.setMaximumSize(1000, 80)

        # Horizontal box containing export level buttons, Create buttons with levels labels
        self.levelContainer = QHBoxLayout(self.horizontalLevelWidget)
        self.levels = ['Practice', 'PCN', 'ICB', 'Region', 'Country']
        for level in self.levels:
            self.cb_level = QPushButton(level)
            self.cb_level.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
            self.cb_level.setMinimumSize(100, 50)
            self.cb_level.setCheckable(True)
            self.levelContainer.addWidget(self.cb_level)
            self.cb_level.clicked.connect(self.LevelSelect)

        # Add Spacing
        self.placeVerticalLayout.addSpacing(25)

        # Filter Group Box
        self.filterGroup = QGroupBox(title='Filter by...')
        self.placeVerticalLayout.addWidget(self.filterGroup)
        self.filterGroup.setMaximumSize(1000, 180)
        self.filterGroup.setStyleSheet('QGroupBox'
                                       '{'
                                       'border: 1px solid black;'
                                       '}'
                                       'font-weight: normal;'
                                       'font-size: 10pt;'
                                       )

        # Filter apply button
        self.applyButton = QPushButton('Apply Filter', parent=self.filterGroup)
        self.applyButton.setGeometry(520, 130, 100, 30)
        self.applyButton.setEnabled(False)
        self.applyButton.setStyleSheet('QPushButton'
                                       '{'
                                       'background-color: limegreen;'
                                       'border-style: outset;'
                                       'border-width: 2px;'
                                       'border-color: limegreen;'
                                       'font-size: 8pt;'
                                       '}'
                                       'QPushButton:hover'
                                       '{'
                                       'background-color: lightgreen'
                                       '}'
                                       'QPushButton:disabled'
                                       '{'
                                       'background-color: darkgrey'
                                       '}'
                                       )
        self.applyButton.clicked.connect(self.ApplyFilter)

        # Filter remove button
        self.removeButton = QPushButton('Remove Filter', parent=self.filterGroup)
        self.removeButton.setGeometry(410, 130, 100, 30)
        self.removeButton.setEnabled(False)
        self.removeButton.setStyleSheet('QPushButton'
                                        '{'
                                        'background-color: red;'
                                        'border-style: outset;'
                                        'border-width: 2px;'
                                        'border-color: red;'
                                        'font-size: 8pt;'
                                        '}'
                                        'QPushButton:hover'
                                        '{'
                                        'background-color: #FC7676'
                                        '}'
                                        'QPushButton:disabled'
                                        '{'
                                        'background-color: darkgrey'
                                        '}'
                                        )
        self.removeButton.clicked.connect(self.RemoveFilter)

        # Horizontal container for buttons and dropdown
        self.filterContainer = QHBoxLayout(self.filterGroup)

        # Vertical container to list filter buttons, create using levels (minus Practice)
        self.filterButtonContainer = QVBoxLayout()
        self.filterButtonContainer.addSpacing(5)
        for level in self.levels[1:]:
            self.filter = QRadioButton(level)
            self.filter.setMaximumSize(100, 100)
            self.filterButtonContainer.addWidget(self.filter)
            self.filter.clicked.connect(self.SelectFilterButton)
            self.filter.setEnabled(False)
        self.filterContainer.addLayout(self.filterButtonContainer)

        # Filter select dropdown box
        self.filterSelectionBox = QComboBox()
        self.filterContainer.addWidget(self.filterSelectionBox)
        self.filterSelectionBox.currentTextChanged.connect(self.RemoveFilter)

        # Add Spacing
        self.placeVerticalLayout.addSpacing(40)

        # Export select group container
        self.selectAreasGroup = QGroupBox('Export data for...')
        self.placeVerticalLayout.addWidget(self.selectAreasGroup)
        self.selectAreasGroup.setMaximumSize(1000, 100)
        self.selectAreasGroup.setGeometry(QRect(3, 400, 640, 80))
        self.selectAreasGroup.setStyleSheet('QGroupBox'
                                            '{'
                                            'border: 1px solid black;'
                                            '}'
                                            'font-weight: normal;'
                                            'font-size: 10pt;'
                                            )

        # Export select dropdown box, also adds practice name to codes for display
        self.selectAreas = QComboBox(self.selectAreasGroup)
        self.practiceCodes = self.mapFrame.PRACTICE_CODE.unique()
        self.practiceDisplay = [
            f'{(self.mapFrame.loc[self.mapFrame["PRACTICE_CODE"] == code, "Practice"]).values[0]} - {code}'
            for code in self.practiceCodes]
        self.practiceDisplay.sort()
        # self.selectAreas.addItems(self.practiceDisplay)
        self.selectAreas.setGeometry(QRect(20, 40, 600, 40))

        # Add spacing
        self.placeVerticalLayout.addSpacing(50)

        # Dataset Settings Group
        self.settingsGroup = QGroupBox('Export settings...')
        self.settingsGroup.setObjectName('MainBox')
        self.grid.addWidget(self.settingsGroup, 0, 3, 1, 2)

        self.settingsVlayout = QVBoxLayout()
        self.settingsGroup.setLayout(self.settingsVlayout)

        # Add Spacing
        self.settingsVlayout.addSpacing(40)

        # Prevalence or Achievement group container
        self.achOrPrevGroup = QGroupBox('Select dataset...', parent=self.settingsGroup)
        self.settingsVlayout.addWidget(self.achOrPrevGroup)
        self.achOrPrevGroup.setStyleSheet('QGroupBox'
                                          '{'
                                          'border: 1px solid black;'
                                          '}'
                                          )

        # H box for prev vs ach radio buttons
        self.achOrPrevVertical = QVBoxLayout(self.achOrPrevGroup)
        self.achOrPrevHorizontal = QHBoxLayout()
        self.achOrPrevHorizontal.setAlignment(Qt.AlignCenter)
        self.achOrPrevVertical.addSpacing(15)
        self.achOrPrevVertical.addLayout(self.achOrPrevHorizontal)

        # Achievement & Prevalence Radio Buttons
        achPrevList = ['Achievement', 'Prevalence']
        for set in achPrevList:
            self.achPrevButton = QRadioButton(set)
            self.achOrPrevHorizontal.addWidget(self.achPrevButton)
            self.achPrevButton.setObjectName('AchPrev')
            self.achPrevButton.clicked.connect(self.SelectDataset)
            if set == self.defaultDataset:
                self.currentDataset = set
                self.achPrevButton.setChecked(True)
            if set == 'Achievement':
                self.achOrPrevHorizontal.addSpacing(100)

        self.achOrPrevVertical.addSpacing(15)

        # Stacked layout settings
        self.settingsStack = QStackedWidget(parent=self.settingsGroup)
        self.settingsVlayout.addWidget(self.settingsStack)

        # Achievement Settings
        self.buttonalignment = Qt.AlignLeft
        self.buttonminsize = 140
        self.maxWidgetHeight = 480
        self.widgetcolor = '#06244F'
        self.achSettingVBox = QVBoxLayout()
        self.achSettingsWidget = QWidget()
        self.achSettingsWidget.setMaximumSize(1000, self.maxWidgetHeight)
        self.achSettingsWidget.setStyleSheet('QWidget'
                                             '{'
                                             'border-radius: 5px;'
                                             f'background-color: {self.widgetcolor};'
                                             '}')
        self.achSettingsWidget.setLayout(self.achSettingVBox)

        # How to handle multiple practices
        self.achPracticeSettingsGroup = QGroupBox('How to group practices...')
        self.achPracticeSettingsGroup.setStyleSheet('QGroupBox'
                                                    '{'
                                                    'border: 0.5px solid grey;'
                                                    'font-size: 8pt;'
                                                    '}'
                                                    'QGroupBox::Title'
                                                    '{'
                                                    'background-color: #06244F;'
                                                    'color: darkgrey;'
                                                    '}')
        self.achPracticeSettings = QHBoxLayout()
        self.achPracticeSettingsGroup.setLayout(self.achPracticeSettings)
        self.achPracticeSettings.setAlignment(self.buttonalignment)
        self.achPracticeSettingButtons = QButtonGroup()
        self.practiceSettings = ['Keep Separate', 'Average', 'Sum']
        for practiceSetting in self.practiceSettings:
            newWid = QRadioButton(practiceSetting)
            newWid.setMinimumSize(self.buttonminsize, 1)
            newWid.setObjectName(f'practices:{practiceSetting}')
            newWid.clicked.connect(self.UpdateAchievementSettings)
            self.achPracticeSettings.addWidget(newWid)
            self.achPracticeSettingButtons.addButton(newWid)
            if practiceSetting == 'Keep Separate':
                newWid.setChecked(True)
                self.achPracticeSetSel = practiceSetting

        # Which data columns to export
        self.achWhichColumnsGroup = QGroupBox('Which columns to include...')
        self.achWhichColumnsGroup.setStyleSheet('QGroupBox'
                                                '{'
                                                'border: 1px solid grey;'
                                                'font-size: 8pt;'
                                                '}'
                                                'QGroupBox::Title'
                                                '{'
                                                'background-color: #06244F;'
                                                'color: darkgrey;'
                                                '}')
        self.achWhichColumns = QHBoxLayout()
        self.achWhichColumnsGroup.setLayout(self.achWhichColumns)
        self.achWhichColumns.setAlignment(self.buttonalignment)
        self.achWhichColumnsButtons = QButtonGroup()
        self.achColumnSettings = ['All Columns', 'All exc. PCAS', 'Score && Percentage', 'Percentage Only']
        for colFunc in self.achColumnSettings:
            newWid = QRadioButton(colFunc)
            newWid.setMinimumSize(self.buttonminsize, 1)
            newWid.setObjectName(f'columns:{colFunc}')
            newWid.clicked.connect(self.UpdateAchievementSettings)
            self.achWhichColumns.addWidget(newWid)
            self.achWhichColumnsButtons.addButton(newWid)
            if colFunc == 'All Columns':
                newWid.setChecked(True)
                self.achColSetSel = colFunc

        # How to group Indicators
        self.achKpiSettingsGroup = QGroupBox('How to group indicator scores...')
        self.achKpiSettingsGroup.setStyleSheet('QGroupBox'
                                               '{'
                                               'border: 1px solid grey;'
                                               'font-size: 8pt;'
                                               '}'
                                               'QGroupBox::Title'
                                               '{'
                                               'background-color: #06244F;'
                                               'color: darkgrey;'
                                               '}')
        self.achKpiSettings = QHBoxLayout()
        self.achKpiSettingsGroup.setLayout(self.achKpiSettings)
        self.achKpiSettings.setAlignment(self.buttonalignment)
        self.achKpiSettingButtons = QButtonGroup()
        self.kpiSettings = ['Keep Separate', 'By Disease Area', 'Total Score']
        for kpi in self.kpiSettings:
            newWid = QRadioButton(kpi)
            newWid.setMinimumSize(self.buttonminsize, 1)
            newWid.setObjectName(f'kpis:{kpi}')
            newWid.clicked.connect(self.UpdateAchievementSettings)
            self.achKpiSettings.addWidget(newWid)
            self.achKpiSettingButtons.addButton(newWid)
            if kpi == 'Keep Separate':
                newWid.setChecked(True)
                self.achKpiSetSel = kpi

        # Which Indicators to Export
        self.achPickKpiSettingsGroup = QGroupBox('Which indicators to include...')
        self.achPickKpiSettingsGroup.setStyleSheet('QGroupBox'
                                                   '{'
                                                   'border: 1px solid grey;'
                                                   'font-size: 8pt;'
                                                   '}'
                                                   'QGroupBox::Title'
                                                   '{'
                                                   'background-color: #06244F;'
                                                   'color: darkgrey;'
                                                   '}')
        self.achPickKpiSettings = QHBoxLayout()
        self.achPickKpiSettingsGroup.setLayout(self.achPickKpiSettings)
        self.achPickKpiSettings.setAlignment(self.buttonalignment)
        self.achPickKpiSettingButtons = QButtonGroup()
        self.kpiPickSettings = ['All Indicators', '2021-22 Indicators Only']
        for pick in self.kpiPickSettings:
            newPick = QRadioButton(pick)
            newPick.setMinimumSize(self.buttonminsize, 1)
            newPick.setObjectName(f'pick:{pick}')
            newPick.clicked.connect(self.UpdateAchievementSettings)
            self.achPickKpiSettings.addWidget(newPick)
            self.achPickKpiSettingButtons.addButton(newPick)
            if pick == 'All Indicators':
                newPick.setChecked(True)
                self.achPickKpiSetSel = pick

        # Add Horizontal Button Layouts to Achievement Vertical Layout
        self.achSettingVBox.addWidget(self.achPracticeSettingsGroup)
        self.achSettingVBox.addWidget(self.achWhichColumnsGroup)
        self.achSettingVBox.addWidget(self.achKpiSettingsGroup)
        self.achSettingVBox.addWidget(self.achPickKpiSettingsGroup)

        # Prevalence Settings
        self.prevSettingVBox = QVBoxLayout()
        self.prevSettingVBox.setAlignment(Qt.AlignTop)
        self.prevSettingsWidget = QWidget()
        self.prevSettingsWidget.setMaximumSize(1000, self.maxWidgetHeight)
        self.prevSettingsWidget.setStyleSheet('QWidget'
                                              '{'
                                              'border-radius: 5px;'
                                              f'background-color: {self.widgetcolor};'
                                              '}')
        self.prevSettingsWidget.setLayout(self.prevSettingVBox)

        self.prevPracticeSettingsGroup = QGroupBox('How to group practices...')
        self.prevPracticeSettingsGroup.setMinimumHeight(97)
        self.prevPracticeSettingsGroup.setStyleSheet('QGroupBox'
                                                     '{'
                                                     'border: 1px solid grey;'
                                                     'font-size: 8pt;'
                                                     '}'
                                                     'QGroupBox::Title'
                                                     '{'
                                                     'background-color: #06244F;'
                                                     'color: darkgrey;'
                                                     '}')
        self.prevPracticeSettingsHBox = QHBoxLayout()
        self.prevPracticeSettingsHBox.setAlignment(self.buttonalignment)
        self.prevPracticeSettingsGroup.setLayout(self.prevPracticeSettingsHBox)
        self.prevPracticeSettingsButtons = QButtonGroup()
        for pracSet in self.practiceSettings[:2]:
            newPrevWid = QRadioButton(pracSet)
            newPrevWid.setMinimumWidth(self.buttonminsize)
            newPrevWid.setObjectName(f'practices:{pracSet}')
            newPrevWid.clicked.connect(self.UpdatePrevalenceSettings)
            self.prevPracticeSettingsHBox.addWidget(newPrevWid)
            self.prevPracticeSettingsButtons.addButton(newPrevWid)
            if pracSet == 'Keep Separate':
                newPrevWid.setChecked(True)
                self.prevPracSetSel = pracSet

        self.prevColumnSettingsGroup = QGroupBox('Which columns to include...')
        self.prevColumnSettingsGroup.setMinimumHeight(98)
        self.prevColumnSettingsGroup.setStyleSheet('QGroupBox'
                                                   '{'
                                                   'border: 1px solid grey;'
                                                   'font-size: 8pt;'
                                                   '}'
                                                   'QGroupBox::Title'
                                                   '{'
                                                   'background-color: #06244F;'
                                                   'color: darkgrey;'
                                                   '}')
        self.prevColumnSettingsHBox = QHBoxLayout()
        self.prevColumnSettingsHBox.setAlignment(self.buttonalignment)
        self.prevColumnSettingsGroup.setLayout(self.prevColumnSettingsHBox)
        self.prevColumnSettingsButtons = QButtonGroup()
        self.prevColumnSettingsList = ['All Columns', 'All exc. List Type', 'Percentage Only']
        for colSet in self.prevColumnSettingsList:
            newPrevButton = QRadioButton(colSet)
            newPrevButton.setMinimumWidth(self.buttonminsize)
            newPrevButton.setObjectName(f'columns:{colSet}')
            newPrevButton.clicked.connect(self.UpdatePrevalenceSettings)
            self.prevColumnSettingsHBox.addWidget(newPrevButton)
            self.prevColumnSettingsButtons.addButton(newPrevButton)
            if colSet == 'All Columns':
                newPrevButton.setChecked(True)
                self.prevColSetSel = colSet

        # Add prev buttons to widget
        self.prevSettingVBox.addWidget(self.prevPracticeSettingsGroup)
        self.prevSettingVBox.addWidget(self.prevColumnSettingsGroup)

        # Add Settings Widgets to Stack
        self.settingsStack.addWidget(self.achSettingsWidget)
        self.settingsStack.addWidget(self.prevSettingsWidget)

        # Export button, connected to export function. Only works with years & area selected
        self.exportButton = QPushButton('Export', parent=self.settingsGroup)
        self.pbar = QProgressBar(self)
        self.exportButton.setObjectName('exportButton')
        self.exportButton.setCheckable(True)
        self.exportButton.pressed.connect(self.PrepareExport)
        self.exportButton.setMinimumSize(200, 60)
        self.exportButton.setMaximumSize(200, 100)
        self.exportHLayout = QHBoxLayout()
        self.settingsVlayout.addLayout(self.exportHLayout)
        self.exportHLayout.setAlignment(Qt.AlignRight)
        self.exportHLayout.addWidget(self.pbar)
        self.exportHLayout.addWidget(self.exportButton)

        self.pbar.hide()

    # Select all functionality for years
    def SelectAllYears(self):
        items_checked = self.year_group.findChildren(QCheckBox)
        sb = self.sender()
        if sb.text() == 'Select all':
            sb.setText('Unselect all')
            for item in items_checked:
                item.setChecked(True)

        elif sb.text() == 'Unselect all':
            sb.setText('Select all')
            for item in items_checked:
                item.setChecked(False)

    # Sets text box entry to current selection
    def YearsOnClick(self):
        txt = ''
        self.currentYears = []
        items_checked = self.year_group.findChildren(QCheckBox)
        for item in items_checked:
            if item.isChecked() and item.objectName() != 'selectAll':
                self.currentYears.append(item.text())
                txt += item.text() + '\n'
        self.date_list.setText('Years Selected: \n \n' + txt)

    # Selects level, deselects all other levels. Also, affects filter functionality
    def LevelSelect(self):
        new_level = self.sender()

        if new_level.isChecked():  # If button is a new level
            self.currentLevel = new_level.text()
            allLevels = self.horizontalLevelWidget.findChildren(QPushButton)

            for level in allLevels:
                if level.text() != new_level.text():
                    level.setChecked(False)

            if self.applyButton.isEnabled():  # Apply button on (i.e. filter not enabled)
                self.DisplayLevels(self.currentLevel)
                if self.currentFilter:
                    if self.levels.index(self.currentLevel) >= self.levels.index(
                            self.currentFilter):
                        self.filterSelectionBox.clear()
                        self.applyButton.setEnabled(False)

            elif not self.applyButton.isEnabled() and self.filterSelectionBox.count() == 0:  # Startup, nothing enabled
                self.DisplayLevels(self.currentLevel)
                self.ActivateFilterButtons()

            elif self.levels.index(self.currentLevel) >= self.levels.index(
                    self.currentFilter):  # If level goes above filter
                self.filterSelectionBox.clear()  # Clear filter box
                self.RemoveFilter(function=True)  # Remove filter, display standard levels

            else:  # Changing level below filter
                self.ApplyFilter()

            self.ActivateFilterButtons()
            self.filterSelectionBox.setEnabled(True)  # Re-enable once disabled

            if self.currentLevel == 'Practice':
                allButtons = self.findChildren(QRadioButton)
                for button in allButtons:
                    if 'practices' in button.objectName():
                        if 'Keep Separate' in button.objectName():
                            button.setChecked(True)
                        button.setEnabled(False)
            else:
                allButtons = self.findChildren(QRadioButton)
                for button in allButtons:
                    if 'practices' in button.objectName():
                        if self.achColSetSel == 'Percentage Only':
                            if 'Sum' not in button.objectName():
                                button.setEnabled(True)
                        else:
                            button.setEnabled(True)

        elif not new_level.isChecked():  # If press is old button being unchecked
            self.selectAreas.clear()
            self.filterSelectionBox.setEnabled(False)
            self.applyButton.setEnabled(False)
            self.removeButton.setEnabled(False)
            self.DeactivateFilterButtons()  # Turn off all buttons

    # Enables all filter buttons beyond current level, removes filter if level changes too high
    def ActivateFilterButtons(self):
        cutoff = self.levels.index(self.currentLevel)
        toActivate = self.levels[cutoff + 1:]
        allRadioButtons = self.filterGroup.findChildren(QRadioButton)
        if self.currentLevel == 'Country':  # Can't filter for country
            self.applyButton.setEnabled(False)

        for button in allRadioButtons:
            if button.text() not in toActivate:  # Deactivate buttons below current level
                button.setAutoExclusive(False)  # If true, at least one button must be selected
                button.setChecked(False)
                button.setAutoExclusive(True)
                button.setEnabled(False)
            else:
                button.setEnabled(True)  # Activate any button above current level

    # Turns off all filter buttons
    def DeactivateFilterButtons(self):
        allRadioButtons = self.filterGroup.findChildren(QRadioButton)
        for button in allRadioButtons:
            button.setEnabled(False)

    # Displays the export options, secondaryFilter and displayList allow filter functionality
    def DisplayLevels(self, selectedLevel=None, displayList=None, secondaryFilter=False):
        if not secondaryFilter:
            level = selectedLevel

            if level != 'Practice':
                filteredMapValues = self.mapFrame[level].unique().tolist()
                filteredMapValues.sort()
                self.selectAreas.clear()
                self.selectAreas.addItems(filteredMapValues)
            elif level == 'Practice':  # Practice separate to enable code/name combo to be shown
                self.selectAreas.clear()
                self.selectAreas.addItems(self.practiceDisplay)

        elif secondaryFilter:
            self.selectAreas.clear()
            self.selectAreas.addItems(displayList)

    # Used when radio button is clicked, removes current filter & turns off remove button
    def SelectFilterButton(self):
        self.FillFilterBox()
        if self.removeButton.isEnabled():
            self.RemoveFilter(function=True)
        self.removeButton.setEnabled(False)
        self.applyButton.setEnabled(True)  # Must be after RemoveFilter to reactivate

    # Uses sender (filter dropdown) to filter mapping frame, sort items, & add to filter box
    def FillFilterBox(self):
        filterBy = self.sender().text()
        self.currentFilter = filterBy
        self.applyButton.setEnabled(True)
        self.filterSelectionBox.clear()
        newItems = self.mapFrame[filterBy].unique().tolist()
        newItems.sort()
        self.filterSelectionBox.addItems(newItems)

    # Sets contents of export box to values allowed by the filter
    def ApplyFilter(self):
        self.applyButton.setEnabled(False)
        self.removeButton.setEnabled(True)
        if self.currentLevel != 'Practice':
            newItems = (self.mapFrame.loc[self.mapFrame[self.currentFilter] == self.filterSelectionBox.currentText(),
                                          self.currentLevel]).unique().tolist()
            newItems.sort()
            self.DisplayLevels(secondaryFilter=True, displayList=newItems)
        else:
            names = (self.mapFrame.loc[self.mapFrame[self.currentFilter] == self.filterSelectionBox.currentText(),
                                       self.currentLevel]).unique().tolist()
            newItems = [
                f'{name} - {(self.mapFrame.loc[self.mapFrame[self.currentLevel] == name, "PRACTICE_CODE"]).values[0]}'
                for name in names]
            newItems.sort()
            self.DisplayLevels(secondaryFilter=True, displayList=newItems)

    # Removes filter from export box, func == True added as function was being given self value so was returning True
    def RemoveFilter(self, function=False):
        if function == True:  # Part of separate filter functionality, not driven by user
            self.removeButton.setEnabled(False)
            self.applyButton.setEnabled(False)  # Deactivate both for case when level > filter
            self.DisplayLevels(self.currentLevel)
        else:
            if self.sender().metaObject().className() == 'QPushButton':  # Activated by remove button
                self.removeButton.setEnabled(False)
                self.applyButton.setEnabled(True)
                self.DisplayLevels(self.currentLevel)

            if self.sender().metaObject().className() == 'QComboBox':  # Activated on textChanged
                if not self.applyButton.isEnabled():
                    self.ApplyFilter()

    # Changes setting stack
    def SelectDataset(self):
        senderText = self.sender().text()
        self.currentDataset = senderText
        if senderText == 'Achievement':
            self.settingsStack.setCurrentIndex(0)
        else:
            self.settingsStack.setCurrentIndex(1)

    # Ach Settings for Export
    def UpdateAchievementSettings(self):
        name = self.sender().objectName()

        if 'columns' in name:
            self.achColSetSel = name.split(':')[1]

            practiceToTurnOff = self.achSettingsWidget.findChild(QRadioButton, 'practices:Sum')
            practiceToTurnOn = self.achSettingsWidget.findChild(QRadioButton, 'practices:Average')
            if 'Percentage Only' in name:
                if self.achPracticeSetSel == 'Sum':
                    practiceToTurnOn.setChecked(True)
                practiceToTurnOff.setEnabled(False)
            elif self.currentLevel is not None:
                if self.currentLevel != 'Practice':
                    practiceToTurnOff.setEnabled(True)
            else:
                practiceToTurnOff.setEnabled(True)

        elif 'practices' in name:
            self.achPracticeSetSel = name.split(':')[1]
            pass

        if 'kpis' in name:
            self.achKpiSetSel = name.split(':')[1]
            columnToTurnOff = self.achSettingsWidget.findChild(QRadioButton, 'columns:All Columns')
            columnToTurnOffToo = self.achSettingsWidget.findChild(QRadioButton, 'columns:All exc. PCAS')
            columnToTurnOn = self.achSettingsWidget.findChild(QRadioButton, 'columns:Score && Percentage')
            if ('Total Score' in name) or ('By Disease' in name):
                if self.achColSetSel == 'All Columns' or self.achColSetSel == 'All exc. PCAS':
                    columnToTurnOn.setChecked(True)
                    self.achColSetSel = columnToTurnOn.objectName().split(':')[1]

                columnToTurnOff.setEnabled(False)
                columnToTurnOffToo.setEnabled(False)
            else:
                columnToTurnOff.setEnabled(True)
                columnToTurnOffToo.setEnabled(True)

        elif 'pick' in name:
            self.achPickKpiSetSel = name.split(':')[1]
            pass

    # Prev Settings for Export
    def UpdatePrevalenceSettings(self):
        name = self.sender().objectName()

        if 'columns' in name:
            self.prevColSetSel = name.split(':')[1]
        elif 'practices' in name:
            self.prevPracSetSel = name.split(':')[1]

    # Opens CSV files
    def ReadCsv(self, name, index=None, header=None):
        frame = None
        frame = pd.read_csv(name, index_col=index, header=header)
        if frame is not None:
            return frame
        else:
            raise LookupError

    # Master filter function
    def MasterFilter(self, dataset, codes, frame=None):
        if dataset == 'Achievement':
            frame = self.achFrame.loc[self.achFrame.index.isin(codes, level=0)]
            frame = self.FilterByAchievementSettings(frame)
        elif dataset == 'Prevalence':
            frame = self.prevFrame.loc[self.prevFrame.index.isin(codes, level=0)]
            frame = self.FilterByPrevalenceSettings(frame)

        if frame.index.nlevels < 2:
            frame = self.ReindexExportFrame(frame)
        if frame.columns.nlevels < 2:
            frame = self.RecolumnExportFrame(frame)

        exportFrame = self.FilterByYears(frame)

        return exportFrame

    # Produces final frame for Achievement
    def FilterByAchievementSettings(self, frame):

        # Which Indicators to include
        if self.achPickKpiSetSel == 'All Indicators':
            pass
        elif self.achPickKpiSetSel == '2021-22 Indicators Only':
            frameCopy = frame.copy()
            for code in frameCopy.columns.get_level_values(0).unique().tolist():
                if (frameCopy.loc[(frameCopy.index.get_level_values(0).unique().tolist()[0:5],
                                   '2021-22'), code]).isna().all().all():
                    frameCopy.drop(code, axis=1, level=0, inplace=True)

            frame = frameCopy.copy()

        # How to group multiple practices
        if self.achPracticeSetSel == 'Keep Separate':
            pass
        elif self.achPracticeSetSel == 'Sum':
            headers = frame.columns.get_level_values(0)

            frame = frame.groupby(level=1).sum(min_count=1)
            frame.loc[:, (headers, 'Percentage Achievement')] = \
                (100 * (frame.loc[:, (headers, 'Score')]).droplevel([1], axis=1) /
                 (frame.loc[:, (headers, 'Max Score')]).droplevel([1], axis=1)).values
            frame.sort_index(ascending=False, inplace=True)

        elif self.achPracticeSetSel == 'Average':
            frame = frame.groupby(level=1).mean()
            frame.sort_index(ascending=False, inplace=True)

        # How to group Indicators
        if self.achKpiSetSel == 'Keep Separate':
            pass
        elif self.achKpiSetSel == 'By Disease Area':

            diseaseAreas = frame.columns.get_level_values(0).unique().tolist()
            diseaseCodes = []
            for disease in diseaseAreas:
                diseaseCode = disease.split('0')[0]
                diseaseCodes.append(diseaseCode)
            diseaseCodes = [*set(diseaseCodes)]
            diseaseCodes.sort()

            frame = frame.drop(['PCAS', 'Register', 'Numerator', 'Denominator'], axis=1, level=1)
            framesToAppend = []
            for kpi in ['Score', 'Max Score', 'Percentage Achievement']:
                newFrame = frame.loc[:, frame.columns.get_level_values(1) == kpi]
                newFrame = newFrame.groupby(lambda x: x.split('0')[0], axis=1, level=0).sum(min_count=1)
                newFrame.columns = pd.MultiIndex.from_product([newFrame.columns.values.tolist(), [kpi]])
                framesToAppend.append(newFrame)

            frame = (pd.concat(framesToAppend, axis=1)).sort_index(axis=1)
            frame.loc[:, frame.columns.get_level_values(1).isin(['Percentage Achievement'])] \
                = (100 * frame.xs('Score', axis=1, level=1) / frame.xs('Max Score', axis=1, level=1)).values
            frame = frame.reindex(['Score', 'Max Score', 'Percentage Achievement'], axis=1, level=1)

        elif self.achKpiSetSel == 'Total Score':
            frame = frame.drop(['PCAS', 'Register', 'Numerator', 'Denominator'], axis=1, level=1)
            frame = frame.groupby(axis=1, level=1).sum(min_count=1)
            frame['Percentage Achievement'] = 100 * frame['Score'] / frame['Max Score']
            frame = frame[['Score', 'Max Score', 'Percentage Achievement']]

        # Which columns to export
        if self.achColSetSel == 'All Columns':
            pass
        elif self.achColSetSel == 'All exc. PCAS':
            frame = frame.drop('PCAS', axis=1, level=1)
        elif self.achColSetSel == 'Score && Percentage' and self.achKpiSetSel != 'Total Score':
            cols = frame.columns.get_level_values(1).isin(
                ['Score', 'Max Score', 'Percentage Achievement'])
            frame = frame.loc[:, cols]
        elif self.achColSetSel == 'Percentage Only':
            if self.achKpiSetSel == 'Total Score':
                frame = frame[['Percentage Achievement']]
            else:
                frame = frame.xs('Percentage Achievement', axis=1, level=1, drop_level=False)

        return frame

    # Produces final frame for Prevalence
    def FilterByPrevalenceSettings(self, frame):

        if self.prevPracSetSel == 'Keep Separate':
            pass
        elif self.prevPracSetSel == 'Average':
            if self.prevColSetSel == 'All Columns':
                tempFrame = frame.copy()
                tempFrame.loc[:, (tempFrame.columns.get_level_values(0), 'Patient List Type')] = np.nan
                tempFrame = tempFrame.groupby(axis=0, level=1).mean()
                tempFrame.loc[:, (frame.columns.get_level_values(0), 'Patient List Type')] = (
                    frame.iloc[0:6].loc[:, (frame.columns.get_level_values(0), 'Patient List Type')]).values
                frame = tempFrame
                frame.sort_index(ascending=False, inplace=True)
            else:
                frame = frame.drop('Patient List Type', axis=1, level=1)
                frame = frame.groupby(axis=0, level=1).mean()
                frame.sort_index(ascending=False, inplace=True)

        if self.prevColSetSel == 'All Columns':
            pass
        elif self.prevColSetSel == 'All exc. List Type':
            frame = frame.drop('Patient List Type', axis=1, level=1)
        elif self.prevColSetSel == 'Percentage Only':
            frame = frame.xs('Percentage Prevalence', axis=1, level=1)

        return frame

    # Filter by selected years
    def FilterByYears(self, frame):
        if frame.index.nlevels > 1:
            frame = frame.loc[frame.index.isin(self.currentYears, level=1)]
        else:
            frame = frame.loc[frame.index.isin(self.currentYears)]
        return frame

    def ReindexExportFrame(self, frame):
        newMulti = pd.MultiIndex.from_product([[self.selectAreas.currentText()], frame.index.values.tolist()])
        frame = frame.set_index(newMulti)
        self.exportReindexed = True
        return frame

    def RecolumnExportFrame(self, frame):
        if self.currentDataset == 'Achievement':
            newColumnMulti = pd.MultiIndex.from_product([['Total'], frame.columns.values.tolist()])
            frame.columns = newColumnMulti
        elif self.currentDataset == 'Prevalence':
            newColumnMulti = pd.MultiIndex.from_product([['Percentage Prevalence'], frame.columns.values.tolist()])
            frame.columns = newColumnMulti

        return frame

    def SetFilePathName(self, originalName, dlg):
        newName = originalName

        if os.path.exists(newName):
            maxExports = 20
            for i in range(maxExports):
                if os.path.exists(newName):
                    newName = f'{originalName.split(".")[0]} ({i + 1}).xlsx'
                else:
                    break

                if i == maxExports - 1:
                    dlg.setWindowTitle('Error')
                    dlg.setText(f'Exceeded download limit ({maxExports}) for {self.selectAreas.currentText()}. '
                                f'Please delete a copy or move from Downloads folder to continue.')
                    dlg.exec()
                    self.exportButton.setEnabled(True)
                    newName = 'Return'

        return newName

    # Exports using export box selection to Excel file into Downloads folder
    def PrepareExport(self):
        dlg = QMessageBox()
        dlg.setWindowIcon(QIcon(os.path.join(basedir, 'Dependencies/Exclamation.png')))
        dlg.setWindowTitle('Error')
        dlg.setStyleSheet('QLabel'
                          '{'
                          'font: normal 10pt Arial;'
                          'text-align: left;'
                          '}')
        dlg.setIcon(QMessageBox.Warning)

        downloadPath = f'{Path.home()}/Downloads'

        if not self.sender().isChecked() and len(self.currentYears) > 0 and len(self.selectAreas.currentText()) > 0:

            self.exportButton.setEnabled(False)

            if self.currentLevel == 'Country':
                if ((self.currentDataset == 'Achievement' and self.achPracticeSetSel == 'Keep Separate') or
                        (self.currentDataset == 'Prevalence' and self.prevPracSetSel == 'Keep Separate')):
                    warning = QMessageBox()
                    warning.setWindowIcon(QIcon(os.path.join(basedir, 'Dependencies/Exclamation.png')))
                    warning.setWindowTitle('Warning: Large Export Size')
                    warning.setStyleSheet('QLabel'
                                          '{'
                                          'font: normal 10pt Arial;'
                                          'text-align: left;'
                                          '}')
                    warning.setIcon(QMessageBox.Warning)
                    warning.setStandardButtons(QMessageBox.Cancel | QMessageBox.Ok)
                    warning.setText(
                        'Warning!\n\nThis export will take some time to complete due to its large filesize.')
                    result = warning.exec()
                    if result == QMessageBox.Ok:
                        pass
                    elif result == QMessageBox.Cancel:
                        self.exportButton.setEnabled(True)
                        return

            codeList = []
            self.exportReindexed = False

            # Populate List of Practice Codes to Include
            if self.currentLevel == 'Practice':
                code = self.selectAreas.currentText()
                codeList.append(code.split('- ')[1])
            else:
                codeList = self.mapFrame.loc[self.mapFrame[self.currentLevel] == self.selectAreas.currentText(),
                                             'PRACTICE_CODE'].values.tolist()

            exportFrame = self.MasterFilter(dataset=self.currentDataset, codes=codeList)

            finalExport = exportFrame.fillna(value='None')

            if not self.exportReindexed:
                finalExport.rename(lambda i: self.mapFrame.loc[self.mapFrame['PRACTICE_CODE'] == i, 'Practice'].item(),
                                   level=0, inplace=True)

            currentSiteName = self.selectAreas.currentText()
            if '/' in currentSiteName:
                currentSiteName = currentSiteName.replace('/', ' ')
            exportName = f'{downloadPath}/{currentSiteName} {self.currentDataset} Export.xlsx'
            sheetName = f'{self.currentDataset} Data'

            # Avoids naming conflicts up to limit
            exportName = self.SetFilePathName(exportName, dlg=dlg)

            if exportName == 'Return':
                self.exportButton.setEnabled(True)
                return

            current_definitions = None
            if self.currentDataset == 'Achievement':
                current_definitions = self.AchievementDefinitions
            elif self.currentDataset == 'Prevalence':
                current_definitions = self.PrevalenceDefinitions

            self.pbar.setValue(0)
            self.pbar.show()

            worker = Worker(fn=self.RunExport, fileName=exportName, sheet=sheetName, dataset=self.currentDataset,
                            definitions=current_definitions, number_of_sites=len(codeList), df=finalExport)

            worker.signals.progress.connect(self.UpdateProgressBar)
            worker.signals.error.connect(self.ExportError)
            worker.signals.finished.connect(lambda: self.ExportComplete(dialog_box=dlg))

            self.threadpool.start(worker)

        elif len(self.currentYears) == 0:
            dlg.setText('Select years to continue')
            dlg.exec()

        elif len(self.selectAreas.currentText()) == 0:
            dlg.setText('Select site(s) to continue')
            dlg.exec()

    def RunExport(self, fileName, sheet, df, dataset, definitions, number_of_sites, progress_callback):
        workbook = None
        target_value = 98

        workbook = xl.Workbook(filename=fileName)
        header_format = workbook.add_format({
            'bold': 0,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': 'white',
            'fg_color': '#3366FF'})
        index_format_1 = workbook.add_format({
            'bold': 0,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': 1,
            'font_color': 'white',
            'fg_color': '#FF9900',})
        index_format_2 = workbook.add_format({
            'bold': 0,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': 1,
            'font_color': 'white',
            'fg_color': '#FF6600'})
        data_format_1 = workbook.add_format({
            'bold': 0,
            'border': 0,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00',
            'fg_color': '#99CCFF'})
        data_format_2 = workbook.add_format({
            'bold': 0,
            'border': 0,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00',
            'fg_color': '#CCFFFF'})
        definitions_index_format_1 = workbook.add_format({
            'bold': 0,
            'border': 1,
            'left': 2,
            'align': 'left',
            'valign': 'vcenter',
            'font_color': 'white',
            'font_size': 14,
            'fg_color': '#FF9900', })
        definitions_index_format_2 = workbook.add_format({
            'bold': 0,
            'border': 1,
            'left': 2,
            'align': 'left',
            'valign': 'vcenter',
            'font_color': 'white',
            'font_size': 14,
            'fg_color': '#FF6600'})
        definitions_data_format_1 = workbook.add_format({
            'bold': 0,
            'border': 0,
            'right': 2,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': 1,
            'indent': 1,
            'font_size': 10,
            'fg_color': '#99CCFF'})
        definitions_data_format_2 = workbook.add_format({
            'bold': 0,
            'border': 0,
            'right': 2,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': 1,
            'indent': 1,
            'font_size': 10,
            'fg_color': '#CCFFFF'})
        definitions_data_format_3 = workbook.add_format({
            'bold': 0,
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': 0,
            'font_size': 14,
            'fg_color': '#CCFFFF'})
        definitions_header_format = workbook.add_format({
            'bold': 0,
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 16,
            'font_color': 'white',
            'fg_color': '#3366FF'})
        table_bottom_format = workbook.add_format({
            'top': 2
        })
        table_end_format = workbook.add_format({
            'left': 2
        })
        sites_format = workbook.add_format({
            'font_size': 9,
            'align': 'center',
            'fg_color': '#EBF1DE'
        })

        definitions_worksheet = workbook.add_worksheet(f'{dataset} Definitions')
        definitions_worksheet.hide_gridlines(2)
        if dataset == 'Achievement':
            definitions_worksheet.set_column(0, 0, 18)
            definitions_worksheet.set_column(1, 1, 24)
            definitions_worksheet.set_column(2, 2, 110)
            definitions_worksheet.set_column(3, 3, 30)
            definitions_worksheet.set_column('E:XFD', None, None, {'hidden': True})
            definitions_worksheet.set_default_row(hide_unused_rows=True)
            definitions_worksheet.set_row(0, None, None, {'hidden': False} )
            definitions_worksheet.set_row(1, 26)
        elif dataset == 'Prevalence':
            definitions_worksheet.set_column(0, 0, 12)
            definitions_worksheet.set_column(1, 1, 24)
            definitions_worksheet.set_column(2, 2, 65)
            definitions_worksheet.set_column(3, 3, 18)
            definitions_worksheet.set_column(4, 4, 35)
            definitions_worksheet.set_column('H:XFD', None, None, {'hidden': True})
            definitions_worksheet.set_default_row(hide_unused_rows=True)
            definitions_worksheet.set_row(0, None, None, {'hidden': False} )
            definitions_worksheet.set_row(1, 26)

            definitions_worksheet.write('E2', 'List Type', definitions_header_format)
            definitions_worksheet.merge_range(2, 4, 3, 4, 'e.g. 06OV is "6 & Over"', definitions_data_format_3)

            definitions_data_format_1.set_font_size(14)
            definitions_data_format_2.set_font_size(14)

        start_row = 1
        definitions_len = definitions.shape[0]
        definitions_worksheet.write(start_row, 1, definitions.iloc[0, 0], definitions_header_format)
        definitions_worksheet.write(start_row, 2, definitions.iloc[0, 1], definitions_header_format)
        row = start_row + 1
        for i in range(definitions_len-1):
            if i%2 == 0:
                definitions_worksheet.write(row + i, 1, definitions.iloc[i+1,0], definitions_index_format_1)
                definitions_worksheet.write(row + i, 2, definitions.iloc[i+1,1], definitions_data_format_1)
            if i%2 == 1:
                definitions_worksheet.write(row + i, 1, definitions.iloc[i+1, 0], definitions_index_format_2)
                definitions_worksheet.write(row + i, 2, definitions.iloc[i+1, 1], definitions_data_format_2)

        definitions_worksheet.write(start_row + definitions_len, 1, None, table_bottom_format)
        definitions_worksheet.write(start_row + definitions_len, 2, None, table_bottom_format)
        definitions_worksheet.write(start_row + definitions_len, 3, 'Author: Laurie Smith')

        worksheet = workbook.add_worksheet(sheet)
        index_values = df.index.get_level_values(0).tolist()
        subindex_values = df.index.get_level_values(1).unique().tolist()
        column_values = df.columns.get_level_values(0).tolist()
        sub_column_values = df.columns.get_level_values(1).tolist()
        numberOfPractices = int(len(index_values) / len(subindex_values))
        numberOfYears = len(subindex_values)

        step = 0
        total_steps = 10*len(index_values) + len(column_values)

        original_value = self.pbar.value()
        current_value = self.pbar.value()

        even_row = False
        new_col_index = True
        current_col = 0
        number_cols = 0
        cols = []

        if numberOfPractices > 1:
            for i in range(len(index_values)):
                step += 10
                actual_value = round((step / total_steps) * (target_value - original_value))
                if actual_value > current_value:
                    current_value = actual_value
                progress_callback.emit(current_value)

                current_subindex_value = i % numberOfYears

                if current_subindex_value == 0:
                    if not even_row:
                        if numberOfYears > 1:
                            worksheet.merge_range(i + 2, 0, i + 2 + numberOfYears - 1, 0,
                                                  index_values[i], index_format_1)
                        else:
                            worksheet.write(i + 2, 0, index_values[i], index_format_1)
                        even_row = True

                    elif even_row:
                        if numberOfYears > 1:
                            worksheet.merge_range(i + 2, 0, i + 2 + numberOfYears - 1, 0,
                                                  index_values[i], index_format_2)
                        else:
                            worksheet.write(i + 2, 0, index_values[i], index_format_2)
                        even_row = False

                if even_row:
                    worksheet.write(i + 2, 1, subindex_values[current_subindex_value], index_format_1)
                elif not even_row:
                    worksheet.write(i + 2, 1, subindex_values[current_subindex_value], index_format_2)

                for j in range(len(column_values)):
                    if even_row:
                        worksheet.write(i + 2, j + 2, df.iloc[i, j], data_format_1)
                    elif not even_row:
                        worksheet.write(i + 2, j + 2, df.iloc[i, j], data_format_2)

                worksheet.write(i+2, len(column_values)+2, None, table_end_format)

            for k in range(len(column_values)):
                step += 1
                actual_value = round((step / total_steps) * (target_value - original_value))
                if actual_value > current_value:
                    current_value = actual_value
                progress_callback.emit(current_value)
                if new_col_index:
                    tempdf = df[column_values[k]]
                    number_cols = tempdf.shape[1]
                    cols = tempdf.columns.tolist()
                    current_col = 0
                    new_col_index = False
                    if number_cols > 1:
                        worksheet.merge_range(0, k + 2, 0, k + 2 + number_cols - 1, column_values[k],
                                              header_format)
                    else:
                        worksheet.write(0, k + 2, column_values[k], header_format)

                worksheet.write(1, k + 2, cols[current_col], header_format)
                current_col += 1
                if current_col == number_cols:
                    new_col_index = True

                worksheet.write(len(index_values)+ 2, k + 2, None, table_bottom_format)

        elif numberOfPractices == 1:
            even_row = True
            if numberOfYears > 1:
                worksheet.merge_range(2, 0, 2 + numberOfYears - 1, 0,
                                      index_values[0], index_format_1)
            else:
                worksheet.write(2, 0, index_values[0], index_format_1)
            for i in range(len(index_values)):
                step += 10
                actual_value = round((step / total_steps) * (target_value - original_value))
                if actual_value > current_value:
                    current_value = actual_value
                progress_callback.emit(current_value)

                if even_row:
                    worksheet.write(i + 2, 1, subindex_values[i], index_format_1)
                    even_row = False
                elif not even_row:
                    worksheet.write(i + 2, 1, subindex_values[i], index_format_2)
                    even_row = True

                for j in range(len(column_values)):
                    if not even_row:
                        worksheet.write(i + 2, j + 2, df.iloc[i, j], data_format_1)
                    if even_row:
                        worksheet.write(i + 2, j + 2, df.iloc[i, j], data_format_2)

                worksheet.write(i + 2, len(column_values) + 2, None, table_end_format)

            for k in range(len(column_values)):
                step += 1
                actual_value = round((step / total_steps) * (target_value - original_value))
                if actual_value > current_value:
                    current_value = actual_value
                progress_callback.emit(current_value)

                if new_col_index:
                    tempdf = df[column_values[k]]
                    number_cols = tempdf.shape[1]
                    cols = tempdf.columns.tolist()
                    current_col = 0
                    new_col_index = False
                    if number_cols > 1:
                        worksheet.merge_range(0, k + 2, 0, k + 2 + number_cols - 1, column_values[k],
                                              header_format)
                    else:
                        worksheet.write(0, k + 2, column_values[k], header_format)

                worksheet.write(1, k + 2, cols[current_col], header_format)
                current_col += 1
                if current_col == number_cols:
                    new_col_index = True

                worksheet.write(len(index_values) + 2, k + 2, None, table_bottom_format)

        for i, string in enumerate(sub_column_values):
            string_width = len(string)
            if string_width < 11:
                string_width = 11
            else:
                string_width += 1
            worksheet.set_column(i+2, i+2, string_width)

        cells_under_image = workbook.add_format({'fg_color': 'white'})
        worksheet.insert_image('A1', os.path.join(basedir, 'Dependencies/unity_logo.png'),
                               {'x_offset': 3, 'y_offset': 3, 'x_scale': 0.15, 'y_scale': 0.15})
        worksheet.write('A1', '', cells_under_image)
        worksheet.write('B2', '', cells_under_image)

        worksheet.write('B1', '# of practices', sites_format)
        worksheet.write('B2', number_of_sites, sites_format)

        worksheet.set_column(0, 0, 15)
        worksheet.set_column(1, 1, 10)
        worksheet.freeze_panes(2, 2)

        worksheet.activate()

        workbook.close()

        current_value = 100
        progress_callback.emit(current_value)

    def ExportComplete(self, dialog_box):
        dialog_box.setWindowTitle('Download Successful')
        dialog_box.setWindowIcon(QIcon(os.path.join(basedir, 'Dependencies/greencheck.png')))
        dialog_box.setIcon(QMessageBox.Information)
        dialog_box.setText('File saved to downloads')
        dialog_box.exec()
        self.pbar.hide()
        self.exportButton.setEnabled(True)

    def ExportError(self):
        print('Error')

    def UpdateProgressBar(self, new_value):
        self.pbar.setValue(new_value)

# Runs application
def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec())


# Calls run app function
window()
