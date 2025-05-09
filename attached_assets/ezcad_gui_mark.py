"""
EZCAD2 GUI Automation Map

This file defines the structure and UI elements of the EZCAD2 application
for automation purposes. It provides a mapping of UI element identifiers 
that can be used with PyWinAuto to automate interactions with EZCAD2.

For more information on how to use these identifiers with PyWinAuto, see:
https://pywinauto.readthedocs.io/en/latest/
"""

# Main application window identifiers
EZCAD_MAIN_WINDOW = {
    'title': 'EZCAD2',
    'class_name': '#32770',
    'control_type': 'Dialog'
}

# Main toolbar buttons
TOOLBAR_BUTTONS = {
    'red': {
        'control_id': 1056,  # This ID may vary, verify with spy tool
        'title': 'Red',
        'class_name': 'Button'
    },
    'mark': {
        'control_id': 1057,  # This ID may vary, verify with spy tool
        'title': 'Mark',
        'class_name': 'Button'
    },
    'test': {
        'control_id': 1058,  # This ID may vary, verify with spy tool
        'title': 'Test',
        'class_name': 'Button'
    },
    'stop': {
        'control_id': 1059,  # This ID may vary, verify with spy tool
        'title': 'Stop',
        'class_name': 'Button'
    },
    'save': {
        'control_id': 1060,  # This ID may vary, verify with spy tool
        'title': 'Save',
        'class_name': 'Button'
    }
}

# Menu identifiers
MENU_ITEMS = {
    'file': {
        'title': 'File',
        'item_id': 0  # Item position in menu bar (0-based index)
    },
    'edit': {
        'title': 'Edit',
        'item_id': 1
    },
    'view': {
        'title': 'View',
        'item_id': 2
    },
    'draw': {
        'title': 'Draw',
        'item_id': 3
    },
    'tools': {
        'title': 'Tools',
        'item_id': 4
    },
    'options': {
        'title': 'Options',
        'item_id': 5
    },
    'help': {
        'title': 'Help',
        'item_id': 6
    }
}

# File menu subitems
FILE_MENU_ITEMS = {
    'new': {
        'title': 'New',
        'item_id': 0
    },
    'open': {
        'title': 'Open',
        'item_id': 1
    },
    'save': {
        'title': 'Save',
        'item_id': 2
    },
    'save_as': {
        'title': 'Save As',
        'item_id': 3
    },
    'exit': {
        'title': 'Exit',
        'item_id': 7  # Position may vary
    }
}

# Dialog identifiers
DIALOGS = {
    'open_file': {
        'title': 'Open',
        'class_name': '#32770'
    },
    'save_file': {
        'title': 'Save As',
        'class_name': '#32770'
    },
    'parameters': {
        'title': 'Parameters',
        'class_name': '#32770'
    },
    'marking_options': {
        'title': 'Marking Options',
        'class_name': '#32770'
    }
}

# Status bar identifiers
STATUS_BAR = {
    'class_name': 'msctls_statusbar32',
    'control_id': 59648  # This ID may vary
}

# Common keyboard shortcuts
KEYBOARD_SHORTCUTS = {
    'new_file': '^n',  # Ctrl+N
    'open_file': '^o',  # Ctrl+O
    'save_file': '^s',  # Ctrl+S
    'undo': '^z',  # Ctrl+Z
    'redo': '^y',  # Ctrl+Y
    'cut': '^x',  # Ctrl+X
    'copy': '^c',  # Ctrl+C
    'paste': '^v',  # Ctrl+V
    'select_all': '^a',  # Ctrl+A
    'delete': '{DELETE}',  # Delete key
    'escape': '{ESC}'  # Escape key
}

# Function key mappings
FUNCTION_KEYS = {
    'f1': '{F1}',  # Help
    'f2': '{F2}',  # Rename
    'f3': '{F3}',  # Find Next
    'f4': '{F4}',  # Repeat
    'f5': '{F5}',  # Refresh
    'f6': '{F6}',  # Next Pane
    'f7': '{F7}',  # Check Spelling
    'f8': '{F8}',  # Run
    'f9': '{F9}',  # Compile
    'f10': '{F10}',  # Menu
    'f11': '{F11}',  # Full Screen
    'f12': '{F12}'  # Save As
}

# Parameter dialog elements
PARAMETER_DIALOG_ELEMENTS = {
    'power': {
        'control_id': 1001,  # This ID may vary, verify with spy tool
        'class_name': 'Edit'
    },
    'speed': {
        'control_id': 1002,  # This ID may vary, verify with spy tool
        'class_name': 'Edit'
    },
    'frequency': {
        'control_id': 1003,  # This ID may vary, verify with spy tool
        'class_name': 'Edit'
    },
    'ok_button': {
        'control_id': 1,  # Common OK button ID
        'class_name': 'Button',
        'title': 'OK'
    },
    'cancel_button': {
        'control_id': 2,  # Common Cancel button ID
        'class_name': 'Button',
        'title': 'Cancel'
    }
}

# Main drawing area
DRAWING_AREA = {
    'class_name': 'AfxFrameOrView42',  # This class name may vary
    'control_id': 59648  # This ID may vary, verify with spy tool
}

# Common automation functions

def get_red_button_spec():
    """
    Get the spec to identify the Red button in EZCAD2 toolbar
    
    Returns:
        dict: Button specification for PyWinAuto
    """
    return TOOLBAR_BUTTONS['red']

def get_mark_button_spec():
    """
    Get the spec to identify the Mark button in EZCAD2 toolbar
    
    Returns:
        dict: Button specification for PyWinAuto
    """
    return TOOLBAR_BUTTONS['mark']

def get_main_window_spec():
    """
    Get the spec to identify the main EZCAD2 window
    
    Returns:
        dict: Window specification for PyWinAuto
    """
    return EZCAD_MAIN_WINDOW

def get_open_file_dialog_spec():
    """
    Get the spec to identify the Open File dialog
    
    Returns:
        dict: Dialog specification for PyWinAuto
    """
    return DIALOGS['open_file']

def get_keyboard_shortcut(action):
    """
    Get the keyboard shortcut for a specific action
    
    Args:
        action (str): Action name as defined in KEYBOARD_SHORTCUTS
        
    Returns:
        str: Keyboard shortcut string for PyWinAuto
    """
    return KEYBOARD_SHORTCUTS.get(action, '')