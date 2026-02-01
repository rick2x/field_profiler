# -*- coding: utf-8 -*-
"""
/***************************************************************************
 FieldProfilerPlugin
                                 A QGIS plugin
 Analyzes and summarizes attribute field values for selected layers.
                              -------------------
        begin                : 2025-05-08
        copyright            : (C) 2025 by Your Name
        email                : your.email@example.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""

from qgis.PyQt.QtGui import QIcon, QKeySequence
from qgis.PyQt.QtWidgets import QAction, QToolBar
from qgis.PyQt.QtCore import Qt, QCoreApplication
import os.path

from .field_profiler_dockwidget import FieldProfilerDockWidget
from qgis.core import Qgis, QgsMessageLog

class FieldProfilerPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.actions = []
        self.menu_name = self.tr("&Field Profiler")
        self.toolbar = None
        self.dockwidget = None
        self.first_run = True

    def tr(self, message):
        return QCoreApplication.translate('FieldProfilerPlugin', message)


    def initGui(self):
        """
        Create the menu entries and toolbar icons for the plugin.
        """
        toolbar_object_name = "FieldProfilerPluginMainToolbar"

        existing_toolbars = self.iface.mainWindow().findChildren(QToolBar, toolbar_object_name)

        if existing_toolbars:
            self.toolbar = existing_toolbars[0]
            self.toolbar.clear()
        else:
            self.toolbar = self.iface.addToolBar(self.tr("Field Profiler"))
            self.toolbar.setObjectName(toolbar_object_name)

        icon_path = os.path.join(self.plugin_dir, 'icon.png')
        action_object_name = "runFieldProfilerActionInstanceUnique"

        self.action = QAction(QIcon(icon_path),
                              self.tr("Run Field Profiler"),
                              self.iface.mainWindow())
        self.action.setObjectName(action_object_name)
        self.action.setStatusTip(self.tr("Analyzes attribute field values"))
        self.action.setWhatsThis(self.tr("Launches the Field Profiler to analyze attribute data."))
        self.action.triggered.connect(self.run)

        self.toolbar.addAction(self.action)
        self.iface.addPluginToMenu(self.menu_name, self.action)
        self.actions.append(self.action)


        self.action.setShortcut(QKeySequence("Ctrl+Alt+Shift+P"))  

        if self.first_run:
            if self.dockwidget is None:
                self.dockwidget = FieldProfilerDockWidget(self.iface, self.iface.mainWindow())
                self.iface.addDockWidget(Qt.RightDockWidgetArea, self.dockwidget)
            self.dockwidget.hide()
            self.first_run = False
    
    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""

        for action in self.actions:
            self.iface.removePluginMenu(self.menu_name, action)
            if self.toolbar:
                self.toolbar.removeAction(action)


        self.actions = []

        if self.dockwidget:
            self.iface.removeDockWidget(self.dockwidget)
            self.dockwidget.deleteLater()
            self.dockwidget = None

        if self.toolbar:
            main_window = self.iface.mainWindow()
            if main_window:
                main_window.removeToolBar(self.toolbar)
            del self.toolbar
            self.toolbar = None


        self.action = None
        self.first_run = True

    def run(self):
        """Run method that loads and shows the dockwidget."""
        if self.dockwidget is None:

            self.dockwidget = FieldProfilerDockWidget(self.iface, self.iface.mainWindow())

            self.iface.addDockWidget(Qt.RightDockWidgetArea, self.dockwidget)
        if self.dockwidget.isVisible():
             if self.dockwidget.isFloating():
                 self.dockwidget.raise_()
                 self.dockwidget.activateWindow()
             else:
                 self.dockwidget.close()
        else:
             self.dockwidget.show()
             self.dockwidget.raise_()
             self.dockwidget.activateWindow()
