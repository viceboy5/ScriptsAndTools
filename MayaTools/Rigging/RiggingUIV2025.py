import maya.cmds as cmds
from PySide6 import QtWidgets, QtCore
import maya.OpenMayaUI as omui
import shiboken6
import re
import sys
script_folder = r"D:\GitRepos\ScriptsAndTools\MayaTools\Rigging"
if script_folder not in sys.path:
    sys.path.append(script_folder)
import CreateRK
import CreateControlsForJoints
import SequentialRenamer
import CreateConstraintsForMatchingJoints
import BrokenFKConstraints

def maya_main_window():
    """
    Get the main Maya window as a QMainWindow instance.
    """
    main_window_ptr = omui.MQtUtil.mainWindow()
    return shiboken6.wrapInstance(int(main_window_ptr), QtWidgets.QWidget)


def create_ui():
    global rigging_ui_window
    rigging_ui_window = QtWidgets.QWidget(parent=maya_main_window())
    rigging_ui_window.setWindowTitle("Rigging Automation")
    rigging_ui_window.setGeometry(200, 200, 400, 400)
    rigging_ui_window.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.WindowStaysOnTopHint)

    layout = QtWidgets.QVBoxLayout()

    color_dropdown = QtWidgets.QComboBox()
    color_dropdown.addItems(["Blue", "Red", "Green"])
    layout.addWidget(color_dropdown)

    create_ctrls_button = QtWidgets.QPushButton("Create Controls")
    create_ctrls_button.clicked.connect(lambda: CreateControlsForJoints.create_controls_clicked(color_dropdown.currentText()))
    layout.addWidget(create_ctrls_button)

    clusters_to_joints_button = QtWidgets.QPushButton("Clusters to Joints")
    clusters_to_joints_button.clicked.connect(create_joint_at_cluster_transforms)  # Make sure this is imported or defined
    layout.addWidget(clusters_to_joints_button)

    rename_input = QtWidgets.QLineEdit()
    rename_input.setPlaceholderText("Enter name with # for numbering (e.g. FK_Jnt_##)")
    layout.addWidget(rename_input)

    sequential_renamer_button = QtWidgets.QPushButton("Sequential Renamer")
    sequential_renamer_button.clicked.connect(lambda: SequentialRenamer.RenameSequentially(rename_input.text()))
    layout.addWidget(sequential_renamer_button)

    create_constraints_button = QtWidgets.QPushButton("Create Constraints for Matching Joints")
    create_constraints_button.clicked.connect(lambda: CreateConstraintsForMatchingJoints.find_and_create_constraints())
    layout.addWidget(create_constraints_button)

    broken_fk_constraints_button = QtWidgets.QPushButton("BrokenFK Constraints")
    broken_fk_constraints_button.clicked.connect(lambda: BrokenFKConstraints.create_broken_fk_constraints())
    layout.addWidget(broken_fk_constraints_button)

    create_rk_button = QtWidgets.QPushButton("Create RK")
    create_rk_button.clicked.connect(lambda: CreateRK.build_rk_system())
    layout.addWidget(create_rk_button)

    rigging_ui_window.setLayout(layout)
    rigging_ui_window.show()

# Run the UI
create_ui()
