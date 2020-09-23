Attribute VB_Name = "basPluginConstants"
Option Explicit

' All source code pertaining to this plug-in architecture is copyright(c) 2001
'  Shawn Bullock.  Author reserves all rights and is granting you limited use
'  to use for educational purposes.  You may only redistribute this in its
'  entirety and without modification unless you include a copy of the original
'  work with this notice.  You may not profit or resell this source code.
'
' Portions of this source code is copyright by Steve McMahon
'  www.vbaccellerator.com. and is subject to the terms and conditions set forth
'  by him.
'
' All source code is provided as-is and without warranty of any kind, implied
'  or not.  Author is not responsible for any damage created by it or misuse of
'  source code and functionality intended by the author or not.
'
Public Const PLUG_ERROR_SUCCESS = 0
Public Const PLUG_ERROR_FILE_NOT_FOUND = 1
Public Const PLUG_ERROR_FAILED = 2
Public Const PLUG_ERROR_UNABLE_TO_REGISTER = 3
Public Const PLUG_ERROR_PARAMETER_REQUIRED = 4
Public Const PLUG_ERROR_NO_VALUES = 5
Public Const PLUG_ERROR_UNABLE_TO_LOAD = 6
Public Const PLUG_ERROR_UNABLE_TO_UNLOAD = 7
Public Const PLUG_ERROR_UNABLE_TO_FIRE_EVENT = 8
Public Const PLUG_ERROR_NO_VALID_INTERFACE = 9
Public Const PLUG_ERROR_LIBRARY_NOT_FOUND = 10
Public Const PLUG_ERROR_NOT_LOADED = 11
Public Const PLUG_ERROR_MODULE_ALREADY_EXISTS = 12
Public Const PLUG_ERROR_BUSY = 13
Public Const PLUG_ERROR_REGISTRY_NOT_FOUND = 14
Public Const PLUG_ERROR_NOT_ACTIVATED = 15
Public Const PLUG_ERROR_UNABLE_TO_DEACTIVATE = 16
Public Const PLUG_ERROR_LISTING_NOT_FOUND = 17
Public Const PLUG_ERROR_INCORRECT_PARAMETER_CRITERIA = 18
Public Const PLUG_ERROR_BAD_VERSION = 19


Public Enum PLUGIN_COLOR_CODES
   PLUG_COLOR_NORMAL = &H0                ' Black  : Normal
   PLUG_COLOR_RED = 255                   ' Red    : Error
   PLUG_COLOR_GREEN = 32512               ' Green  : Undefined
   PLUG_COLOR_BLUE = 16711680             ' Blue   : Undefined
End Enum

Public Enum PLUGIN_STATUS_MODE
   PLUG_STATUS_UNKNOWN = 0
   PLUG_STATUS_ACTIVE = 1
   PLUG_STATUS_INACTIVE = 2
   PLUG_STATUS_LOADING = 3
   PLUG_STATUS_ERROR_LOADING = 4
   PLUG_STATUS_ERROR_UNLOADING = 5
   PLUG_STATUS_ERROR_ACTIVATING = 6
   PLUG_STATUS_ERROR_DEACTIVATING = 7
   PLUG_STATUS_UNLOADING = 8
   PLUG_STATUS_BAD_VERSION = 9
   PLUG_STATUS_LOADED = 10
   PLUG_STATUS_OLD_HOST = 11
   PLUG_STATUS_MODULE_NOT_LOADED = 12
   PLUG_STATUS_FAILED_MISERABLY = 13
End Enum

Public Const PLUG_REGISTRY_NAME = 1
Public Const PLUG_REGISTRY_INTERFACE = 2
Public Const PLUG_REGISTRY_PATH = 3
Public Const PLUG_REGISTRY_DESCRIPTION = 4
Public Const PLUG_REGISTRY_DISPLAY = 5
Public Const PLUG_REGISTRY_STARTUP = 6
Public Const PLUG_REGISTRY_SUPPORTS = 7


