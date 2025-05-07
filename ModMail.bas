Attribute VB_Name = "ModMail"
'**********************************************************
'   The DNS & MXQuery code in this module was adapted from
'   Gregg Housh's MX.OCX code. Many thanks to Gregg for
'   his fine work.
'**********************************************************
'Option Explicit

'========================================================================
' Copyright 1999 - Digital Press, John Rhoton
'
' This program has been written to illustrate the Internet Mail protocols.
' It is provided free of charge and unconditionally.  However, it is not
' intended for production use, and therefore without warranty or any
' implication of support.
'
' You can find an explanation of the concepts behind this code in
' the book:  Programmer's Guide to Internet Mail by John Rhoton,
' Digital Press 1999.  ISBN: 1-55558-212-5.
'
' For ordering information please see http://www.amazon.com or
' you can order directly with http://www.bh.com/digitalpress.
'
'========================================================================

Declare Function lstrcpy& Lib "Kernel32" (ByVal dststring As Any, ByVal srcstring As Any)

Declare Sub hmemcpy Lib "Kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbBytes&)

Global Status As Integer
Global CRLF As String
Global CallSocket As InputSocketDescriptor

Type InputSocketDescriptor
  Family As Integer
  Port As Integer
  Address As Long
  Fill As String * 8
End Type

Type OutputSocketDescriptor
  Family As Integer
  Port As Integer
  Address As Long
  Fill As String * 30
End Type


Type WSAdata_type
  wVersion As Integer
  wHighVersion As Integer
  szDescription As String * 257
  szSystemStatus As String * 129
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpVendorInfo As String * 200
End Type
Global WSAdata As WSAdata_type

Type NameHostEntry
  Name As Long
  Aliases As Long
  AddressType As Integer
  Length As Integer
  AddressPointerPointer As Long
  Buffer As String * 100
End Type

Type sockopt_bool_type         'Used for setting socket options
  val As Integer
End Type
'Global sockopt_bool As sockopt_bool_type

  'These are the Win specific calls which use messages
Declare Function WSAStartup Lib "wsock32.dll" (ByVal a As Integer, B As WSAdata_type) As Integer
Declare Function WSACleanup Lib "wsock32.dll" () As Integer
Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal sock As Integer, ByVal hwnd As Long, ByVal msg As Integer, ByVal event2 As Long) As Integer
Declare Function WSAGetLastError Lib "wsock32.dll" () As Integer

  'Winsock calls in VB format
Declare Function gethostbyname Lib "wsock32.dll" (ByVal HostName As String) As Long
Declare Function gethostbyaddr Lib "wsock32.dll" (HostAddress As Long, ByVal lenaddress As Integer, ByVal pftype As Integer) As Long
Declare Function inet_ntoa Lib "wsock32.dll" (ByVal iaddr As Long) As Long
Declare Function htons Lib "wsock32.dll" (ByVal a As Integer) As Integer

Declare Function socket Lib "wsock32.dll" (ByVal af As Integer, ByVal typesock As Integer, ByVal protocol As Integer) As Integer
Declare Function bind Lib "wsock32.dll" (ByVal s As Integer, addr As InputSocketDescriptor, ByVal namelen As Integer) As Integer
Declare Function connect Lib "wsock32.dll" (ByVal sock As Integer, sockstruct As InputSocketDescriptor, ByVal structlen As Integer) As Integer
Declare Function send Lib "wsock32.dll" (ByVal sock As Integer, ByVal msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function recv Lib "wsock32.dll" (ByVal sock As Integer, ByVal msg As String, ByVal msglen As Integer, ByVal flag As Integer) As Integer
Declare Function closesocket Lib "wsock32.dll" (ByVal sock As Integer) As Integer
Declare Function setsockopt Lib "wsock32.dll" (ByVal sock As Integer, ByVal level As Integer, ByVal optname As Integer, optval As sockopt_bool_type, ByVal optlen As Integer) As Integer

Declare Function listen Lib "wsock32.dll" (ByVal sn As Integer, ByVal blog As Integer) As Integer
Declare Function accept Lib "wsock32.dll" (ByVal sn As Integer, saddr As OutputSocketDescriptor, namelen As Integer) As Integer


'Some constants declarations
Global Const SOCKET_ERROR = -1
Global Const INVALID_SOCKET = -1

Global Const SOCK_STREAM = 1
Global Const AF_INET = 2
Global Const PF_INET = 2

Global Const IPPROTO_TCP = 6
Global Const SOL_SOCKET = &HFFFF
Global Const SO_DEBUG = &H1
Global Const SO_ACCEPTCONN = &H2
Global Const SO_REUSEADDR = &H4
Global Const SO_KEEPALIVE = &H8
Global Const SO_DONTROUTE = &H10
Global Const SO_BROADCAST = &H20
Global Const SO_USELOOPBACK = &H40
Global Const SO_LINGER = &H80
Global Const SO_OOBINLINE = &H100
Global Const SO_DONTLINGER = &HFF7F

Global Const AF_UNSPEC = 0

Global Const FD_READ = 1
Global Const FD_WRITE = 2
Global Const FD_OOB = 4
Global Const FD_ACCEPT = 8
Global Const FD_CONNECT = &H10
Global Const FD_CLOSE = &H20



Global Const CF_LINK = &HBF00
Global Const CF_TEXT = 1
Global Const CF_BITMAP = 2
Global Const CF_METAFILE = 3
Global Const CF_DIB = 8
Global Const CF_PALETTE = 9

' DragOver
Global Const ENTER = 0
Global Const LEAVE = 1
Global Const OVER = 2

' Drag (controls)
Global Const Cancel = 0
Global Const BEGIN_DRAG = 1
Global Const END_DRAG = 2

' Show parameters
Global Const MODAL = 1
Global Const MODELESS = 0

' Arrange Method
' for MDI Forms
Global Const CASCADE = 0
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3

'ZOrder Method
Global Const BRINGTOFRONT = 0
Global Const SENDTOBACK = 1

' Key Codes
Global Const KEY_LBUTTON = &H1
Global Const KEY_RBUTTON = &H2
Global Const KEY_CANCEL = &H3
Global Const KEY_MBUTTON = &H4    ' NOT contiguous with L & RBUTTON
Global Const KEY_BACK = &H8
Global Const KEY_TAB = &H9
Global Const KEY_CLEAR = &HC
Global Const KEY_RETURN = &HD
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_MENU = &H12
Global Const KEY_PAUSE = &H13
Global Const KEY_CAPITAL = &H14
Global Const KEY_ESCAPE = &H1B
Global Const KEY_SPACE = &H20
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_END = &H23
Global Const KEY_HOME = &H24
Global Const KEY_LEFT = &H25
Global Const KEY_UP = &H26
Global Const KEY_RIGHT = &H27
Global Const KEY_DOWN = &H28
Global Const KEY_SELECT = &H29
Global Const KEY_PRINT = &H2A
Global Const KEY_EXECUTE = &H2B
Global Const KEY_SNAPSHOT = &H2C
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_HELP = &H2F

' KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'

Global Const KEY_NUMPAD0 = &H60
Global Const KEY_NUMPAD1 = &H61
Global Const KEY_NUMPAD2 = &H62
Global Const KEY_NUMPAD3 = &H63
Global Const KEY_NUMPAD4 = &H64
Global Const KEY_NUMPAD5 = &H65
Global Const KEY_NUMPAD6 = &H66
Global Const KEY_NUMPAD7 = &H67
Global Const KEY_NUMPAD8 = &H68
Global Const KEY_NUMPAD9 = &H69
Global Const KEY_MULTIPLY = &H6A
Global Const KEY_ADD = &H6B
Global Const KEY_SEPARATOR = &H6C
Global Const KEY_SUBTRACT = &H6D
Global Const KEY_DECIMAL = &H6E
Global Const KEY_DIVIDE = &H6F
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_F13 = &H7C
Global Const KEY_F14 = &H7D
Global Const KEY_F15 = &H7E
Global Const KEY_F16 = &H7F

Global Const KEY_NUMLOCK = &H90

' Variant VarType tags

Global Const V_EMPTY = 0
Global Const V_NULL = 1
Global Const V_INTEGER = 2
Global Const V_LONG = 3
Global Const V_SINGLE = 4
Global Const V_DOUBLE = 5
Global Const V_CURRENCY = 6
Global Const V_DATE = 7
Global Const V_STRING = 8


' Event Parameters

' ErrNum (LinkError)
Global Const WRONG_FORMAT = 1
Global Const DDE_SOURCE_CLOSED = 6
Global Const TOO_MANY_LINKS = 7
Global Const DATA_TRANSFER_FAILED = 8

' QueryUnload
Global Const FORM_CONTROLMENU = 0
Global Const FORM_CODE = 1
Global Const APP_WINDOWS = 2
Global Const APP_TASKMANAGER = 3
Global Const FORM_MDIFORM = 4

' Properties

' Colors
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF

' System Colors
Global Const SCROLL_BARS = &H80000000           ' Scroll-bars gray area.
Global Const DESKTOP = &H80000001               ' Desktop.
Global Const ACTIVE_TITLE_BAR = &H80000002      ' Active window caption.
Global Const INACTIVE_TITLE_BAR = &H80000003    ' Inactive window caption.
Global Const MENU_BAR = &H80000004              ' Menu background.
Global Const WINDOW_BACKGROUND = &H80000005     ' Window background.
Global Const WINDOW_FRAME = &H80000006          ' Window frame.
Global Const MENU_TEXT = &H80000007             ' Text in menus.
Global Const WINDOW_TEXT = &H80000008           ' Text in windows.
Global Const TITLE_BAR_TEXT = &H80000009        ' Text in caption, size box, scroll-bar arrow box..
Global Const ACTIVE_BORDER = &H8000000A         ' Active window border.
Global Const INACTIVE_BORDER = &H8000000B       ' Inactive window border.
Global Const APPLICATION_WORKSPACE = &H8000000C ' Background color of multiple document interface (MDI) applications.
Global Const HIGHLIGHT = &H8000000D             ' Items selected item in a control.
Global Const HIGHLIGHT_TEXT = &H8000000E        ' Text of item selected in a control.
Global Const BUTTON_FACE = &H8000000F           ' Face shading on command buttons.
Global Const BUTTON_SHADOW = &H80000010         ' Edge shading on command buttons.
Global Const GRAY_TEXT = &H80000011             ' Grayed (disabled) text.  This color is set to 0 if the current display driver does not support a solid gray color.
Global Const BUTTON_TEXT = &H80000012           ' Text on push buttons.

' Enumerated Types

' Align (picture box)
Global Const NONE = 0
Global Const ALIGN_TOP = 1
Global Const ALIGN_BOTTOM = 2

' Alignment
Global Const LEFT_JUSTIFY = 0  ' 0 - Left Justify
Global Const RIGHT_JUSTIFY = 1 ' 1 - Right Justify
Global Const CENTER = 2        ' 2 - Center

' BorderStyle (form)
'Global Const NONE = 0          ' 0 - None
Global Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Global Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Global Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

' BorderStyle (Shape and Line)
'Global Const TRANSPARENT = 0    '0 - Transparent
'Global Const SOLID = 1          '1 - Solid
'Global Const DASH = 2         ' 2 - Dash
'Global Const DOT = 3          ' 3 - Dot
'Global Const DASH_DOT = 4     ' 4 - Dash-Dot
'Global Const DASH_DOT_DOT = 5 ' 5 - Dash-Dot-Dot
'Global Const INSIDE_SOLID = 6 ' 6 - Inside Solid

' MousePointer
Global Const DEFAULT = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

' DragMode
Global Const MANUAL = 0    ' 0 - Manual
Global Const AUTOMATIC = 1 ' 1 - Automatic

' DrawMode
Global Const BLACKNESS = 1      ' 1 - Blackness
Global Const NOT_MERGE_PEN = 2  ' 2 - Not Merge Pen
Global Const MASK_NOT_PEN = 3   ' 3 - Mask Not Pen
Global Const NOT_COPY_PEN = 4   ' 4 - Not Copy Pen
Global Const MASK_PEN_NOT = 5   ' 5 - Mask Pen Not
Global Const INVERT = 6         ' 6 - Invert
Global Const XOR_PEN = 7        ' 7 - Xor Pen
Global Const NOT_MASK_PEN = 8   ' 8 - Not Mask Pen
Global Const MASK_PEN = 9       ' 9 - Mask Pen
Global Const NOT_XOR_PEN = 10   ' 10 - Not Xor Pen
Global Const NOP = 11           ' 11 - Nop
Global Const MERGE_NOT_PEN = 12 ' 12 - Merge Not Pen
Global Const COPY_PEN = 13      ' 13 - Copy Pen
Global Const MERGE_PEN_NOT = 14 ' 14 - Merge Pen Not
Global Const MERGE_PEN = 15     ' 15 - Merge Pen
Global Const WHITENESS = 16     ' 16 - Whiteness

' DrawStyle
Global Const SOLID = 0        ' 0 - Solid
Global Const DASH = 1         ' 1 - Dash
Global Const DOT = 2          ' 2 - Dot
Global Const DASH_DOT = 3     ' 3 - Dash-Dot
Global Const DASH_DOT_DOT = 4 ' 4 - Dash-Dot-Dot
Global Const INVISIBLE = 5    ' 5 - Invisible
Global Const INSIDE_SOLID = 6 ' 6 - Inside Solid

' FillStyle
' Global Const SOLID = 0           ' 0 - Solid
Global Const TRANSPARENT = 1       ' 1 - Transparent
Global Const HORIZONTAL_LINE = 2   ' 2 - Horizontal Line
Global Const VERTICAL_LINE = 3     ' 3 - Vertical Line
Global Const UPWARD_DIAGONAL = 4   ' 4 - Upward Diagonal
Global Const DOWNWARD_DIAGONAL = 5 ' 5 - Downward Diagonal
Global Const CROSS = 6             ' 6 - Cross
Global Const DIAGONAL_CROSS = 7    ' 7 - Diagonal Cross

' LinkMode (forms and controls)
' Global Const NONE = 0         ' 0 - None
Global Const LINK_SOURCE = 1    ' 1 - Source (forms only)
Global Const LINK_AUTOMATIC = 1 ' 1 - Automatic (controls only)
Global Const LINK_MANUAL = 2    ' 2 - Manual (controls only)
Global Const LINK_NOTIFY = 3    ' 3 - Notify (controls only)

' LinkMode (kept for VB1.0 compatibility, use new constants instead)
Global Const HOT = 1    ' 1 - Hot (controls only)
'Global Const SERVER = 1 ' 1 - Server (forms only)
Global Const COLD = 2   ' 2 - Cold (controls only)


' ScaleMode
Global Const USER = 0        ' 0 - User
Global Const TWIPS = 1       ' 1 - Twip
Global Const POINTS = 2      ' 2 - Point
Global Const PIXELS = 3      ' 3 - Pixel
Global Const CHARACTERS = 4  ' 4 - Character
Global Const INCHES = 5      ' 5 - Inch
Global Const MILLIMETERS = 6 ' 6 - Millimeter
Global Const CENTIMETERS = 7 ' 7 - Centimeter

' ScrollBar
' Global Const NONE     = 0 ' 0 - None
Global Const HORIZONTAL = 1 ' 1 - Horizontal
Global Const VERTICAL = 2   ' 2 - Vertical
Global Const BOTH = 3       ' 3 - Both

' Shape
Global Const SHAPE_RECTANGLE = 0
Global Const SHAPE_SQUARE = 1
Global Const SHAPE_OVAL = 2
Global Const SHAPE_CIRCLE = 3
Global Const SHAPE_ROUNDED_RECTANGLE = 4
Global Const SHAPE_ROUNDED_SQUARE = 5

' WindowState
Global Const NORMAL = 0    ' 0 - Normal
Global Const MINIMIZED = 1 ' 1 - Minimized
Global Const MAXIMIZED = 2 ' 2 - Maximized

' Check Value
Global Const UNCHECKED = 0 ' 0 - Unchecked
Global Const CHECKED = 1   ' 1 - Checked
Global Const GRAYED = 2    ' 2 - Grayed

' Shift parameter masks
Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4

' Button parameter masks
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4

' Function Parameters
' form1.label1.caption= parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal

' form1.label1.caption= return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

' SetAttr, Dir, GetAttr functions
Global Const ATTR_NORMAL = 0
Global Const ATTR_READONLY = 1
Global Const ATTR_HIDDEN = 2
Global Const ATTR_SYSTEM = 4
Global Const ATTR_VOLUME = 8
Global Const ATTR_DIRECTORY = 16
Global Const ATTR_ARCHIVE = 32

'Grid
'ColAlignment,FixedAlignment Properties
Global Const GRID_ALIGNLEFT = 0
Global Const GRID_ALIGNRIGHT = 1
Global Const GRID_ALIGNCENTER = 2

'Fillstyle Property
Global Const GRID_SINGLE = 0
Global Const GRID_REPEAT = 1


'Data control
'Error event Response arguments
Global Const DATA_ERRCONTINUE = 0
Global Const DATA_ERRDISPLAY = 1

'Editmode property values
Global Const DATA_EDITNONE = 0
Global Const DATA_EDITMODE = 1
Global Const DATA_EDITADD = 2

' Options property values
Global Const DATA_DENYWRITE = &H1
Global Const DATA_DENYREAD = &H2
Global Const DATA_READONLY = &H4
Global Const DATA_APPENDONLY = &H8
Global Const DATA_INCONSISTENT = &H10
Global Const DATA_CONSISTENT = &H20
Global Const DATA_SQLPASSTHROUGH = &H40

'Validate event Action arguments
Global Const DATA_ACTIONCANCEL = 0
Global Const DATA_ACTIONMOVEFIRST = 1
Global Const DATA_ACTIONMOVEPREVIOUS = 2
Global Const DATA_ACTIONMOVENEXT = 3
Global Const DATA_ACTIONMOVELAST = 4
Global Const DATA_ACTIONADDNEW = 5
Global Const DATA_ACTIONUPDATE = 6
Global Const DATA_ACTIONDELETE = 7
Global Const DATA_ACTIONFIND = 8
Global Const DATA_ACTIONBOOKMARK = 9
Global Const DATA_ACTIONCLOSE = 10
Global Const DATA_ACTIONUNLOAD = 11


'OLE Client Control
'Actions
Global Const OLE_CREATE_EMBED = 0
Global Const OLE_CREATE_NEW = 0           'from ole1 control
Global Const OLE_CREATE_LINK = 1
Global Const OLE_CREATE_FROM_FILE = 1     'from ole1 control
Global Const OLE_COPY = 4
Global Const OLE_PASTE = 5
Global Const OLE_UPDATE = 6
Global Const OLE_ACTIVATE = 7
Global Const OLE_CLOSE = 9
Global Const OLE_DELETE = 10
Global Const OLE_SAVE_TO_FILE = 11
Global Const OLE_READ_FROM_FILE = 12
Global Const OLE_INSERT_OBJ_DLG = 14
Global Const OLE_PASTE_SPECIAL_DLG = 15
Global Const OLE_FETCH_VERBS = 17
Global Const OLE_SAVE_TO_OLE1FILE = 18

'OLEType
Global Const OLE_LINKED = 0
Global Const OLE_EMBEDDED = 1
Global Const OLE_NONE = 3

'OLETypeAllowed
Global Const OLE_EITHER = 2

'UpdateOptions
Global Const OLE_AUTOMATIC = 0
Global Const OLE_FROZEN = 1
Global Const OLE_MANUAL = 2

'AutoActivate modes
'Note that OLE_ACTIVATE_GETFOCUS only applies to objects that
'support "inside-out" activation.  See related Verb notes below.
Global Const OLE_ACTIVATE_MANUAL = 0
Global Const OLE_ACTIVATE_GETFOCUS = 1
Global Const OLE_ACTIVATE_DOUBLECLICK = 2

'SizeModes
Global Const OLE_SIZE_CLIP = 0
Global Const OLE_SIZE_STRETCH = 1
Global Const OLE_SIZE_AUTOSIZE = 2

'DisplayTypes
Global Const OLE_DISPLAY_CONTENT = 0
Global Const OLE_DISPLAY_ICON = 1

'Update Event Constants
Global Const OLE_CHANGED = 0
Global Const OLE_SAVED = 1
Global Const OLE_CLOSED = 2
Global Const OLE_RENAMED = 3

'Special Verb Values
Global Const VERB_PRIMARY = 0
Global Const VERB_SHOW = -1
Global Const VERB_OPEN = -2
Global Const VERB_HIDE = -3
Global Const VERB_INPLACEUIACTIVATE = -4
Global Const VERB_INPLACEACTIVATE = -5
'The last two verbs are for objects that support "inside-out" activation,
'meaning they can be edited in-place, and that they support being left
'in-place-active even when the input focus moves to another control or form.
'These objects actually have 2 levels of being active.  "InPlace Active"
'means that the object is ready for the user to click inside it and start
'working with it.  "In-Place UI-Active" means that, in addition, if the object
'has any other UI associated with it, such as floating palette windows,
'that those windows are visible and ready for use.  Any number of objects
'can be "In-Place Active" at a time, although only one can be
'"InPlace UI-Active".

'You can cause an object to move to either one of states programmatically by
'setting the Verb property to the appropriate verb and setting
'Action=OLE_ACTIVATE.

'Also, if you set AutoActivate = OLE_ACTIVATE_GETFOCUS, the server will
'automatically be put into "InPlace UI-Active" state when the user clicks
'on or tabs into the control.

'VerbFlag Bit Masks
Global Const VERBFLAG_GRAYED = &H1
Global Const VERBFLAG_DISABLED = &H2
Global Const VERBFLAG_CHECKED = &H8
Global Const VERBFLAG_SEPARATOR = &H800

'MiscFlag Bits - Or these together as desired for special behaviors

'MEMSTORAGE causes the control to use memory to store the object while
'           it is loaded.  This is faster than the default (disk-tempfile),
'           but can consume a lot of memory for objects whose data takes
'           up a lot of space, such as the bitmap for a paint program.
Global Const OLE_MISCFLAG_MEMSTORAGE = &H1

'DISABLEINPLACE overrides the control's default behavior of allowing
'           in-place activation for objects that support it.  If you
'           are having problems activating an object inplace, you can
'           force it to always activate in a separate window by setting this
'           bit
Global Const OLE_MISCFLAG_DISABLEINPLACE = &H2

'Common Dialog Control
'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'File Open/Save Dialog Flags
Global Const OFN_READONLY = &H1&
Global Const OFN_OVERWRITEPROMPT = &H2&
Global Const OFN_HIDEREADONLY = &H4&
Global Const OFN_NOCHANGEDIR = &H8&
Global Const OFN_SHOWHELP = &H10&
Global Const OFN_NOVALIDATE = &H100&
Global Const OFN_ALLOWMULTISELECT = &H200&
Global Const OFN_EXTENSIONDIFFERENT = &H400&
Global Const OFN_PATHMUSTEXIST = &H800&
Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_CREATEPROMPT = &H2000&
Global Const OFN_SHAREAWARE = &H4000&
Global Const OFN_NOREADONLYRETURN = &H8000&

'Color Dialog Flags
Global Const CC_RGBINIT = &H1&
Global Const CC_FULLOPEN = &H2&
Global Const CC_PREVENTFULLOPEN = &H4&
Global Const CC_SHOWHELP = &H8&

'Fonts Dialog Flags
Global Const CF_SCREENFONTS = &H1&
Global Const CF_PRINTERFONTS = &H2&
Global Const CF_BOTH = &H3&
Global Const CF_SHOWHELP = &H4&
Global Const CF_INITTOLOGFONTSTRUCT = &H40&
Global Const CF_USESTYLE = &H80&
Global Const CF_EFFECTS = &H100&
Global Const CF_APPLY = &H200&
Global Const CF_ANSIONLY = &H400&
Global Const CF_NOVECTORFONTS = &H800&
Global Const CF_NOSIMULATIONS = &H1000&
Global Const CF_LIMITSIZE = &H2000&
Global Const CF_FIXEDPITCHONLY = &H4000&
Global Const CF_WYSIWYG = &H8000&         'must also have CF_SCREENFONTS & CF_PRINTERFONTS
Global Const CF_FORCEFONTEXIST = &H10000
Global Const CF_SCALABLEONLY = &H20000
Global Const CF_TTONLY = &H40000
Global Const CF_NOFACESEL = &H80000
Global Const CF_NOSTYLESEL = &H100000
Global Const CF_NOSIZESEL = &H200000

'Printer Dialog Flags
Global Const PD_ALLPAGES = &H0&
Global Const PD_SELECTION = &H1&
Global Const PD_PAGENUMS = &H2&
Global Const PD_NOSELECTION = &H4&
Global Const PD_NOPAGENUMS = &H8&
Global Const PD_COLLATE = &H10&
Global Const PD_PRINTTOFILE = &H20&
Global Const PD_PRINTSETUP = &H40&
Global Const PD_NOWARNING = &H80&
Global Const PD_RETURNDC = &H100&
Global Const PD_RETURNIC = &H200&
Global Const PD_RETURNDEFAULT = &H400&
Global Const PD_SHOWHELP = &H800&
Global Const PD_USEDEVMODECOPIES = &H40000
Global Const PD_DISABLEPRINTTOFILE = &H80000
Global Const PD_HIDEPRINTTOFILE = &H100000

'Help Constants
Global Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Global Const HELP_QUIT = &H2              'Terminate help
Global Const HELP_INDEX = &H3             'Display index
Global Const HELP_CONTENTS = &H3
Global Const HELP_HELPONHELP = &H4        'Display help on using help
Global Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Global Const HELP_SETCONTENTS = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_KEY = &H101             'Display topic for keyword in offabData
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105      'call the search engine in winhelp

'Error Constants
Global Const CDERR_DIALOGFAILURE = -32768

Global Const CDERR_GENERALCODES = &H7FFF
Global Const CDERR_STRUCTSIZE = &H7FFE
Global Const CDERR_INITIALIZATION = &H7FFD
Global Const CDERR_NOTEMPLATE = &H7FFC
Global Const CDERR_NOHINSTANCE = &H7FFB
Global Const CDERR_LOADSTRFAILURE = &H7FFA
Global Const CDERR_FINDRESFAILURE = &H7FF9
Global Const CDERR_LOADRESFAILURE = &H7FF8
Global Const CDERR_LOCKRESFAILURE = &H7FF7
Global Const CDERR_MEMALLOCFAILURE = &H7FF6
Global Const CDERR_MEMLOCKFAILURE = &H7FF5
Global Const CDERR_NOHOOK = &H7FF4

'Added for CMDIALOG.VBX
Global Const CDERR_CANCEL = &H7FF3
Global Const CDERR_NODLL = &H7FF2
Global Const CDERR_ERRPROC = &H7FF1
Global Const CDERR_ALLOC = &H7FF0
Global Const CDERR_HELP = &H7FEF

Global Const PDERR_PRINTERCODES = &H6FFF
Global Const PDERR_SETUPFAILURE = &H6FFE
Global Const PDERR_PARSEFAILURE = &H6FFD
Global Const PDERR_RETDEFFAILURE = &H6FFC
Global Const PDERR_LOADDRVFAILURE = &H6FFB
Global Const PDERR_GETDEVMODEFAIL = &H6FFA
Global Const PDERR_INITFAILURE = &H6FF9
Global Const PDERR_NODEVICES = &H6FF8
Global Const PDERR_NODEFAULTPRN = &H6FF7
Global Const PDERR_DNDMMISMATCH = &H6FF6
Global Const PDERR_CREATEICFAILURE = &H6FF5
Global Const PDERR_PRINTERNOTFOUND = &H6FF4

Global Const CFERR_CHOOSEFONTCODES = &H5FFF
Global Const CFERR_NOFONTS = &H5FFE

Global Const FNERR_FILENAMECODES = &H4FFF
Global Const FNERR_SUBCLASSFAILURE = &H4FFE
Global Const FNERR_INVALIDFILENAME = &H4FFD
Global Const FNERR_BUFFERTOOSMALL = &H4FFC

Global Const FRERR_FINDREPLACECODES = &H3FFF
Global Const CCERR_CHOOSECOLORCODES = &H2FFF


'---------------------------------------------------------
'      Table of Contents for Visual Basic Professional
'
'       1.  3-D Controls
'           (Frame/Panel/Option/Check/Command/Group Push)
'       2.  Animated Button
'       3.  Gauge Control
'       4.  Graph Control Section
'       5.  Key Status Control
'       6.  Spin Button
'       7.  MCI Control (Multimedia)
'       8.  Masked Edit Control
'       9.  Comm Control
'       10. Outline Control
'---------------------------------------------------------


'-------------------------------------------------------------------
'3D Controls
'-------------------------------------------------------------------
'Alignment (Check Box)
Global Const SSCB_TEXT_RIGHT = 0         '0 - Text to the right
Global Const SSCB_TEXT_LEFT = 1          '1 - Text to the left

'Alignment (Option Button)
Global Const SSOB_TEXT_RIGHT = 0         '0 - Text to the right
Global Const SSOB_TEXT_LEFT = 1          '1 - Text to the left

'Alignment (Frame)
Global Const SSFR_LEFT_JUSTIFY = 0       '0 - Left justify text
Global Const SSFR_RIGHT_JUSTIFY = 1      '1 - Right justify text
Global Const SSFR_CENTER = 2             '2 - Center text

'Alignment (Panel)
Global Const SSPN_LEFT_TOP = 0           '0 - Text to left and top
Global Const SSPN_LEFT_MIDDLE = 1        '1 - Text to left and middle
Global Const SSPN_LEFT_BOTTOM = 2        '2 - Text to left and bottom
Global Const SSPN_RIGHT_TOP = 3          '3 - Text to right and top
Global Const SSPN_RIGHT_MIDDLE = 4       '4 - Text to right and middle
Global Const SSPN_RIGHT_BOTTOM = 5       '5 - Text to right and bottom
Global Const SSPN_CENTER_TOP = 6         '6 - Text to center and top
Global Const SSPN_CENTER_MIDDLE = 7      '7 - Text to center and middle
Global Const SSPN_CENTER_BOTTOM = 8      '8 - Text to center and bottom

'Autosize (Command Button)
Global Const SS_AUTOSIZE_NONE = 0        '0 - No Autosizing
Global Const SSPB_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Global Const SSPB_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Ribbon Button)
'Global Const SS_AUTOSIZE_NONE      = 0  '0 - No Autosizing
Global Const SSRI_AUTOSIZE_PICTOBUT = 1  '0 - Autosize Picture to Button
Global Const SSRI_AUTOSIZE_BUTTOPIC = 2  '0 - Autosize Button to Picture

'Autosize (Panel)
'Global Const SS_AUTOSIZE_NONE    = 0    '0 - No Autosizing
Global Const SSPN_AUTOSIZE_WIDTH = 1     '1 - Autosize Panel width to Caption
Global Const SSPN_AUTOSIZE_HEIGHT = 2    '2 - Autosize Panel height to Caption
Global Const SSPN_AUTOSIZE_CHILD = 3     '3 - Autosize Child to Panel

'BevelInner (Panel)
Global Const SS_BEVELINNER_NONE = 0      '0 - No Inner Bevel
Global Const SS_BEVELINNER_INSET = 1     '1 - Inset Inner Bevel
Global Const SS_BEVELINNER_RAISED = 2    '2 - Raised Inner Bevel

'BevelOuter (Panel)
Global Const SS_BEVELOUTER_NONE = 0      '0 - No Outer Bevel
Global Const SS_BEVELOUTER_INSET = 1     '1 - Inset Outer Bevel
Global Const SS_BEVELOUTER_RAISED = 2    '2 - Raised Outer Bevel

'FloodType (Panel)
Global Const SS_FLOODTYPE_NONE = 0       '0 - No flood
Global Const SS_FLOODTYPE_L_TO_R = 1     '1 - Left to light
Global Const SS_FLOODTYPE_R_TO_L = 2     '2 - Right to left
Global Const SS_FLOODTYPE_T_TO_B = 3     '3 - Top to bottom
Global Const SS_FLOODTYPE_B_TO_T = 4     '4 - Bottom to top
Global Const SS_FLOODTYPE_CIRCLE = 5     '5 - Widening circle

'Font3D (Panel, Command Button, Option Button, Check Box, Frame)
Global Const SS_FONT3D_NONE = 0          '0 - No 3-D text
Global Const SS_FONT3D_RAISED_LIGHT = 1  '1 - Raised with light shading
Global Const SS_FONT3D_RAISED_HEAVY = 2  '2 - Raised with heavy shading
Global Const SS_FONT3D_INSET_LIGHT = 3   '3 - Inset with light shading
Global Const SS_FONT3D_INSET_HEAVY = 4   '4 - Inset with heavy shading

'PictureDnChange (Ribbon Button)
Global Const SS_PICDN_NOCHANGE = 0       '0 - Use 'Up'bitmap with no change
Global Const SS_PICDN_DITHER = 1         '1 - Dither 'Up'bitmap
Global Const SS_PICDN_INVERT = 2         '2 - Invert 'Up'bitmap

'ShadowColor (Panel, Frame)
Global Const SS_SHADOW_DARKGREY = 0      '0 - Dark grey shadow
Global Const SS_SHADOW_BLACK = 1         '1 - Black shadow

'ShadowStyle (Frame)
Global Const SS_SHADOW_INSET = 0         '0 - Shadow inset
Global Const SS_SHADOW_RAISED = 1        '1 - Shadow raised


'---------------------------------------
'Animated Button
'---------------------------------------
'Cycle property
Global Const ANI_ANIMATED = 0
Global Const ANI_MULTISTATE = 1
Global Const ANI_TWO_STATE = 2

'Click Filter property
Global Const ANI_ANYWHERE = 0
Global Const ANI_IMAGE_AND_TEXT = 1
Global Const ANI_IMAGE = 2
Global Const ANI_TEXT = 3

'PicDrawMode Property
Global Const ANI_XPOS_YPOS = 0
Global Const ANI_AUTOSIZE = 1
Global Const ANI_STRETCH = 2

'SpecialOp Property
Global Const ANI_CLICK = 1

'TextPosition Property
Global Const ANI_CENTER = 0
Global Const ANI_LEFT = 1
Global Const ANI_RIGHT = 2
Global Const ANI_BOTTON = 3
Global Const ANI_TOP = 4


'---------------------------------------
'GAUGE
'---------------------------------------
'Style Property
Global Const GAUGE_HORIZ = 0
Global Const GAUGE_VERT = 1
Global Const GAUGE_SEMI = 2
Global Const GAUGE_FULL = 3


'----------------------------------------
'Graph Control
'----------------------------------------
'General
Global Const G_NONE = 0
Global Const G_DEFAULT = 0

Global Const G_OFF = 0
Global Const G_ON = 1

Global Const G_MONO = 0
Global Const G_COLOR = 1

'Graph Types
Global Const G_PIE2D = 1
Global Const G_PIE3D = 2
Global Const G_BAR2D = 3
Global Const G_BAR3D = 4
Global Const G_GANTT = 5
Global Const G_LINE = 6
Global Const G_LOGLIN = 7
Global Const G_AREA = 8
Global Const G_SCATTER = 9
Global Const G_POLAR = 10
Global Const G_HLC = 11

'Colors
Global Const G_BLACK = 0
Global Const G_BLUE = 1
Global Const G_GREEN = 2
Global Const G_CYAN = 3
Global Const G_RED = 4
Global Const G_MAGENTA = 5
Global Const G_BROWN = 6
Global Const G_LIGHT_GRAY = 7
Global Const G_DARK_GRAY = 8
Global Const G_LIGHT_BLUE = 9
Global Const G_LIGHT_GREEN = 10
Global Const G_LIGHT_CYAN = 11
Global Const G_LIGHT_RED = 12
Global Const G_LIGHT_MAGENTA = 13
Global Const G_YELLOW = 14
Global Const G_WHITE = 15
Global Const G_AUTOBW = 16

'Patterns
Global Const G_SOLID = 0
Global Const G_HOLLOW = 1
Global Const G_HATCH1 = 2
Global Const G_HATCH2 = 3
Global Const G_HATCH3 = 4
Global Const G_HATCH4 = 5
Global Const G_HATCH5 = 6
Global Const G_HATCH6 = 7
Global Const G_BITMAP1 = 16
Global Const G_BITMAP2 = 17
Global Const G_BITMAP3 = 18
Global Const G_BITMAP4 = 19
Global Const G_BITMAP5 = 20
Global Const G_BITMAP6 = 21
Global Const G_BITMAP7 = 22
Global Const G_BITMAP8 = 23
Global Const G_BITMAP9 = 24
Global Const G_BITMAP10 = 25
Global Const G_BITMAP11 = 26
Global Const G_BITMAP12 = 27
Global Const G_BITMAP13 = 28
Global Const G_BITMAP14 = 29
Global Const G_BITMAP15 = 30
Global Const G_BITMAP16 = 31

'Symbols
Global Const G_CROSS_PLUS = 0
Global Const G_CROSS_TIMES = 1
Global Const G_TRIANGLE_UP = 2
Global Const G_SOLID_TRIANGLE_UP = 3
Global Const G_TRIANGLE_DOWN = 4
Global Const G_SOLID_TRIANGLE_DOWN = 5
Global Const G_SQUARE = 6
Global Const G_SOLID_SQUARE = 7
Global Const G_DIAMOND = 8
Global Const G_SOLID_DIAMOND = 9

'Line Styles
'Global Const G_SOLID = 0
Global Const G_DASH = 1
Global Const G_DOT = 2
Global Const G_DASHDOT = 3
Global Const G_DASHDOTDOT = 4

'Grids
Global Const G_HORIZONTAL = 1
Global Const G_VERTICAL = 2

'Statistics
Global Const G_MEAN = 1
Global Const G_MIN_MAX = 2
Global Const G_STD_DEV = 4
Global Const G_BEST_FIT = 8

'Data Arrays
Global Const G_GRAPH_DATA = 1
Global Const G_COLOR_DATA = 2
Global Const G_EXTRA_DATA = 3
Global Const G_LABEL_TEXT = 4
Global Const G_LEGEND_TEXT = 5
Global Const G_PATTERN_DATA = 6
Global Const G_SYMBOL_DATA = 7
Global Const G_XPOS_DATA = 8
Global Const G_ALL_DATA = 9

'Draw Mode
Global Const G_NO_ACTION = 0
Global Const G_CLEAR = 1
Global Const G_DRAW = 2
Global Const G_BLIT = 3
Global Const G_COPY = 4
Global Const G_PRINT = 5
Global Const G_WRITE = 6

'Print Options
Global Const G_BORDER = 2

'Pie Chart Options             '
Global Const G_NO_LINES = 1
Global Const G_COLORED = 2
Global Const G_PERCENTS = 4

'Bar Chart Options             '
'Global Const G_HORIZONTAL = 1
Global Const G_STACKED = 2
Global Const G_PERCENTAGE = 4
Global Const G_Z_CLUSTERED = 6

'Gantt Chart Options           '
Global Const G_SPACED_BARS = 1

'Line/Polar Chart Options      '
Global Const G_SYMBOLS = 1
Global Const G_STICKS = 2
Global Const G_LINES = 4

'Area Chart Options            '
Global Const G_ABSOLUTE = 1
Global Const G_PERCENT = 2

'HLC Chart Options             '
Global Const G_NO_CLOSE = 1
Global Const G_NO_HIGH_LOW = 2


'---------------------------------------
'Key Status Control
'---------------------------------------
'Style
Global Const KEYSTAT_CAPSLOCK = 0
Global Const KEYSTAT_NUMLOCK = 1
Global Const KEYSTAT_INSERT = 2
Global Const KEYSTAT_SCROLLLOCK = 3


'---------------------------------------
'MCI Control (Multimedia)
'---------------------------------------
'NOTE:
'Please use the updated Multimedia constants
'in the WINMMSYS.TXT file from the \VB\WINAPI
'subdirectory.

'Mode Property
'Global Const MCI_MODE_NOT_OPEN = 11
'Global Const MCI_MODE_STOP = 12
'Global Const MCI_MODE_PLAY = 13
'Global Const MCI_MODE_RECORD = 14
'Global Const MCI_MODE_SEEK = 15
'Global Const MCI_MODE_PAUSE = 16
'Global Const MCI_MODE_READY = 17

'NotifyValue Property
'Global Const MCI_NOTIFY_SUCCESSFUL = 1
'Global Const MCI_NOTIFY_SUPERSEDED = 2
'Global Const MCI_ABORTED = 4
'Global Const MCI_FAILURE = 8

'Orientation Property
'Global Const MCI_ORIENT_HORZ = 0
'Global Const MCI_ORIENT_VERT = 1

'RecordMode Porperty
'Global Const MCI_RECORD_INSERT = 0
'Global Const MCI_RECORD_OVERWRITE = 1

'TimeFormat Property
'Global Const MCI_FORMAT_MILLISECONDS = 0
'Global Const MCI_FORMAT_HMS = 1
'Global Const MCI_FORMAT_MSF = 2
'Global Const MCI_FORMAT_FRAMES = 3
'Global Const MCI_FORMAT_SMPTE_24 = 4
'Global Const MCI_FORMAT_SMPTE_25 = 5
'Global Const MCI_FORMAT_SMPTE_30 = 6
'Global Const MCI_FORMAT_SMPTE_30DROP = 7
'Global Const MCI_FORMAT_BYTES = 8
'Global Const MCI_FORMAT_SAMPLES = 9
'Global Const MCI_FORMAT_TMSF = 10


'---------------------------------------
'Spin Button
'---------------------------------------
'SpinOrientation
Global Const SPIN_VERTICAL = 0
Global Const SPIN_HORIZONTAL = 1


'---------------------------------------
'Masked Edit Control
'---------------------------------------
'ClipMode
Global Const ME_INCLIT = 0
Global Const ME_EXCLIT = 1


'---------------------------------------
'Comm Control
'---------------------------------------
'Handshaking
Global Const MSCOMM_HANDSHAKE_NONE = 0
Global Const MSCOMM_HANDSHAKE_XONXOFF = 1
Global Const MSCOMM_HANDSHAKE_RTS = 2
Global Const MSCOMM_HANDSHAKE_RTSXONXOFF = 3

'Event constants
Global Const MSCOMM_EV_SEND = 1
Global Const MSCOMM_EV_RECEIVE = 2
Global Const MSCOMM_EV_CTS = 3
Global Const MSCOMM_EV_DSR = 4
Global Const MSCOMM_EV_CD = 5
Global Const MSCOMM_EV_RING = 6
Global Const MSCOMM_EV_EOF = 7

'Error code constants
Global Const MSCOMM_ER_BREAK = 1001
Global Const MSCOMM_ER_CTSTO = 1002
Global Const MSCOMM_ER_DSRTO = 1003
Global Const MSCOMM_ER_FRAME = 1004
Global Const MSCOMM_ER_OVERRUN = 1006
Global Const MSCOMM_ER_CDTO = 1007
Global Const MSCOMM_ER_RXOVER = 1008
Global Const MSCOMM_ER_RXPARITY = 1009
Global Const MSCOMM_ER_TXFULL = 1010


'---------------------------------------
' MAPI SESSION CONTROL CONSTANTS
'---------------------------------------
'Action
Global Const SESSION_SIGNON = 1
Global Const SESSION_SIGNOFF = 2


'---------------------------------------
' MAPI MESSAGE CONTROL CONSTANTS
'---------------------------------------
'Action
Global Const MESSAGE_FETCH = 1             ' Load all messages from message store
Global Const MESSAGE_SENDDLG = 2           ' Send mail bring up default mapi dialog
Global Const MESSAGE_SEND = 3              ' Send mail without default mapi dialog
Global Const MESSAGE_SAVEMSG = 4           ' Save message in the compose buffer
Global Const MESSAGE_COPY = 5              ' Copy current message to compose buffer
Global Const MESSAGE_COMPOSE = 6           ' Initialize compose buffer (previous
                                           ' data is lost
Global Const MESSAGE_REPLY = 7             ' Fill Compose buffer as REPLY
Global Const MESSAGE_REPLYALL = 8          ' Fill Compose buffer as REPLY ALL
Global Const MESSAGE_FORWARD = 9           ' Fill Compose buffer as FORWARD
Global Const MESSAGE_DELETE = 10           ' Delete current message
Global Const MESSAGE_SHOWADBOOK = 11       ' Show Address book
Global Const MESSAGE_SHOWDETAILS = 12      ' Show details of the current recipient
Global Const MESSAGE_RESOLVENAME = 13      ' Resolve the display name of the recipient
Global Const RECIPIENT_DELETE = 14            ' Fill Compose buffer as FORWARD
Global Const ATTACHMENT_DELETE = 15          ' Delete current message


'---------------------------------------
'  ERROR CONSTANT DECLARATIONS (MAPI CONTROLS)
'---------------------------------------
Global Const SUCCESS_SUCCESS = 32000
Global Const MAPI_USER_ABORT = 32001
Global Const MAPI_E_FAILURE = 32002
Global Const MAPI_E_LOGIN_FAILURE = 32003
Global Const MAPI_E_DISK_FULL = 32004
Global Const MAPI_E_INSUFFICIENT_MEMORY = 32005
Global Const MAPI_E_ACCESS_DENIED = 32006
Global Const MAPI_E_TOO_MANY_SESSIONS = 32008
Global Const MAPI_E_TOO_MANY_FILES = 32009
Global Const MAPI_E_TOO_MANY_RECIPIENTS = 32010
Global Const MAPI_E_ATTACHMENT_NOT_FOUND = 32011
Global Const MAPI_E_ATTACHMENT_OPEN_FAILURE = 32012
Global Const MAPI_E_ATTACHMENT_WRITE_FAILURE = 32013
Global Const MAPI_E_UNKNOWN_RECIPIENT = 32014
Global Const MAPI_E_BAD_RECIPTYPE = 32015
Global Const MAPI_E_NO_MESSAGES = 32016
Global Const MAPI_E_INVALID_MESSAGE = 32017
Global Const MAPI_E_TEXT_TOO_LARGE = 32018
Global Const MAPI_E_INVALID_SESSION = 32019
Global Const MAPI_E_TYPE_NOT_SUPPORTED = 32020
Global Const MAPI_E_AMBIGUOUS_RECIPIENT = 32021
Global Const MAPI_E_MESSAGE_IN_USE = 32022
Global Const MAPI_E_NETWORK_FAILURE = 32023
Global Const MAPI_E_INVALID_EDITFIELDS = 32024
Global Const MAPI_E_INVALID_RECIPS = 32025
Global Const MAPI_E_NOT_SUPPORTED = 32026

Global Const CONTROL_E_SESSION_EXISTS = 32050
Global Const CONTROL_E_INVALID_BUFFER = 32051
Global Const CONTROL_E_INVALID_READ_BUFFER_ACTION = 32052
Global Const CONTROL_E_NO_SESSION = 32053
Global Const CONTROL_E_INVALID_RECIPIENT = 32054
Global Const CONTROL_E_INVALID_COMPOSE_BUFFER_ACTION = 32055
Global Const CONTROL_E_FAILURE = 32056
Global Const CONTROL_E_NO_RECIPIENTS = 32057
Global Const CONTROL_E_NO_ATTACHMENTS = 32058


'---------------------------------------
'  MISCELLANEOUS GLOBAL CONSTANT DECLARATIONS (MAPI CONTROLS)
'---------------------------------------
Global Const RECIPTYPE_ORIG = 0
Global Const RECIPTYPE_TO = 1
Global Const RECIPTYPE_CC = 2
Global Const RECIPTYPE_BCC = 3

Global Const ATTACHTYPE_DATA = 0
Global Const ATTACHTYPE_EOLE = 1
Global Const ATTACHTYPE_SOLE = 2


'-------------------------------------------------
'  Outline
'-------------------------------------------------
' PictureType
Global Const MSOUTLINE_PICTURE_CLOSED = 0
Global Const MSOUTLINE_PICTURE_OPEN = 1
Global Const MSOUTLINE_PICTURE_LEAF = 2

'Outline Control Error Constants
Global Const MSOUTLINE_BADPICFORMAT = 32000
Global Const MSOUTLINE_BADINDENTATION = 32001
Global Const MSOUTLINE_MEM = 32002
Global Const MSOUTLINE_PARENTNOTEXPANDED = 32003


' Dun
'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RasSetEntryProperties Lib "rasapi32.dll" Alias "RasSetEntryPropertiesA" _
        (ByVal lpszPhonebook As String, ByVal lpszEntry As String, lpRasEntry As Any, ByVal dwEntryInfoSize As Long, _
        lpbDeviceInfo As Any, ByVal dwDeviceInfoSize As Long) As Long
Private Declare Function iRasEnumDevices Lib "rasapi32.dll" Alias "RasEnumDevicesA" (lpRasDevInfo As Any, _
    lpcb As Long, lpcDevices As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Sub iCopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function RasGetEntryProperties Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" (ByVal lpszPhonebook As String, _
        ByVal lpszEntry As String, lpRasEntry As Any, lpdwEntryInfoSize As Long, lpbDeviceInfo As Any, lpdwDeviceInfoSize As Long) As Long
        
Public mlConnectionNumber As Long
Private Type RAS_ENTRIES
    dwSize As Long
    szEntryname(256) As Byte
End Type

Private Type RASIPADDR
    a As Byte
    B As Byte
    c As Byte
    d As Byte
End Type

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Const RAS_MaxDeviceType = 16
Private Const RAS_MaxDeviceName = 128
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0&

Public Type RASDEVINFO
    dwSize As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Public Enum RasEntryOptions
   RASEO_UseCountryAndAreaCodes = &H1
   RASEO_SpecificIpAddr = &H2
   RASEO_SpecificNameServers = &H4
   RASEO_IpHeaderCompression = &H8
   RASEO_RemoteDefaultGateway = &H10
   RASEO_DisableLcpExtensions = &H20
   RASEO_TerminalBeforeDial = &H40
   RASEO_TerminalAfterDial = &H80
   RASEO_ModemLights = &H100
   RASEO_SwCompression = &H200
   RASEO_RequireEncryptedPw = &H400
   RASEO_RequireMsEncryptedPw = &H800
   RASEO_RequireDataEncryption = &H1000
   RASEO_NetworkLogon = &H2000
   RASEO_UseLogonCredentials = &H4000
   RASEO_PromoteAlternates = &H8000
   RASEO_SecureLocalFiles = &H10000
   RASEO_RequireEAP = &H20000
   RASEO_RequirePAP = &H40000
   RASEO_RequireSPAP = &H80000
   RASEO_Custom = &H100000
   RASEO_PreviewPhoneNumber = &H200000
   RASEO_SharedPhoneNumbers = &H800000
   RASEO_PreviewUserPw = &H1000000
   RASEO_PreviewDomain = &H2000000
   RASEO_ShowDialingProgress = &H4000000
   RASEO_RequireCHAP = &H8000000
   RASEO_RequireMsCHAP = &H10000000
   RASEO_RequireMsCHAP2 = &H20000000
   RASEO_RequireW95MSCHAP = &H40000000
   RASEO_CustomScript = &H80000000
End Enum

Public Enum RASNetProtocols
   RASNP_NetBEUI = &H1
   RASNP_Ipx = &H2
   RASNP_Ip = &H4
End Enum

Public Enum RasFramingProtocols
   RASFP_Ppp = &H1
   RASFP_Slip = &H2
   RASFP_Ras = &H4
End Enum


Public Type VBRasEntry
   options As RasEntryOptions
   CountryID As Long
   CountryCode As Long
   AreaCode As String
   LocalPhoneNumber As String
   AlternateNumbers As String
   ipAddr As RASIPADDR
   ipAddrDns As RASIPADDR
   ipAddrDnsAlt As RASIPADDR
   ipAddrWins As RASIPADDR
   ipAddrWinsAlt As RASIPADDR
   FrameSize As Long
   fNetProtocols As RASNetProtocols
   FramingProtocol As RasFramingProtocols
   ScriptName As String
   AutodialDll As String
   AutodialFunc As String
   DeviceType As String
   DeviceName As String
   X25PadType As String
   X25Address As String
   X25Facilities As String
   X25UserData As String
   Channels As Long
   NT4En_SubEntries As Long
   NT4En_DialMode As Long
   NT4En_DialExtraPercent As Long
   NT4En_DialExtraSampleSeconds As Long
   NT4En_HangUpExtraPercent As Long
   NT4En_HangUpExtraSampleSeconds As Long
   NT4En_IdleDisconnectSeconds As Long
   Win2000_Type As Long
   Win2000_EncryptionType As Long
   Win2000_CustomAuthKey As Long
   Win2000_guidId(0 To 15) As Byte
   Win2000_CustomDialDll As String
   Win2000_VpnStrategy As Long
End Type
' winsock
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = &H2
'   Flags for InternetDial - must not conflict with InternetAutodial flags
'                          as they are valid here also.

'   for Dial and Hangup functions
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwnd As Long, ByVal sConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long
    '       Returns   ERROR_SUCCESS if successfull or one of the following error codes
    '                 ERROR_INVALID_PARAMETER - one or more parameters are incorrect
    '                 ERROR_NO_CONNECTION - There is a problem with the dial-up connection
    '                 ERROR_USER_DISCONNECTION - The user clicked either the work offline or cancel button on the dialog box
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long

' Registry access
'Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

' misc
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Public Function StartDUN(hwnd As Long) As Long
    Dim plResult As Long, n As String
    
    n = GetDeviceName
    If Len(n) = 0 Then
        CreateInternetEntry
        n = "KTUNESCO"
    End If
    If mlConnectionNumber <> 0 And ActiveConnection() Then
        StartDUN = -1   'already issued a connection
    Else
        plResult = InternetDial(hwnd, n, INTERNET_AUTODIAL_FORCE_UNATTENDED, mlConnectionNumber, 0&)
        StartDUN = plResult
    End If
End Function

Private Function ActiveConnection() As Boolean
   Dim hKey As Long
   Dim lpData As Long
   Dim nSizeData As Long
   
  'function checks registry for an active connection
   Const sSubKey = "System\CurrentControlSet\Services\RemoteAccess"
   Const sKeyValue = "Remote Connection"
   'default false
   ActiveConnection = False
   If RegOpenKey(HKEY_LOCAL_MACHINE, sSubKey, hKey) = ERROR_SUCCESS Then
      lpData = 0&
      nSizeData = Len(lpData)
      If RegQueryValueEx(hKey, sKeyValue, 0&, 0&, lpData, nSizeData) = ERROR_SUCCESS Then
         ActiveConnection = lpData <> 0
      End If
      Call RegCloseKey(hKey)
   End If
End Function

Public Function HangUp() As Long
    If mlConnectionNumber = 0 Then
        HangUp = -1 'no connection from this module
    Else
        HangUp = InternetHangUp(mlConnectionNumber, 0&)
        mlConnectionNumber = 0&
    End If
End Function

Private Sub CreateInternetEntry()
'make the DUN called DUN_NAME
'Will work if the selected device in the combo box is a modem
'The Connection Created is a "dummy" one
'If you want a real one, change the below parameters
'to the what you need, or add text boxes to a form
'so the user can enter them him/herself
Dim dname As String
dname = GetDeviceName
If Len(dname) = 0 Then Exit Sub
Dim typVBRasEntry As VBRasEntry
typVBRasEntry.AreaCode = ""
typVBRasEntry.AutodialFunc = 1
typVBRasEntry.CountryCode = "1"
typVBRasEntry.CountryID = "1"
typVBRasEntry.DeviceName = dname
typVBRasEntry.DeviceType = "Modem"
typVBRasEntry.fNetProtocols = RASNP_Ip
typVBRasEntry.FramingProtocol = RASFP_Ppp
typVBRasEntry.options = RASEO_SwCompression + RASEO_IpHeaderCompression + RASEO_RemoteDefaultGateway _
+ RASEO_SpecificNameServers + RASEO_PreviewPhoneNumber + RASEO_PreviewUserPw + RASEO_NetworkLogon _
+ RASEO_ShowDialingProgress


typVBRasEntry.ipAddrDns.a = "206"
typVBRasEntry.ipAddrDns.B = "211"
typVBRasEntry.ipAddrDns.c = "214"
typVBRasEntry.ipAddrDns.d = "206"
typVBRasEntry.ipAddrDnsAlt.a = "212"
typVBRasEntry.ipAddrDnsAlt.B = "200"
typVBRasEntry.ipAddrDnsAlt.c = "200"
typVBRasEntry.ipAddrDnsAlt.d = "200"
typVBRasEntry.ipAddrWins.a = "0"
typVBRasEntry.ipAddrWins.B = "0"
typVBRasEntry.ipAddrWins.c = "0"
typVBRasEntry.ipAddrWins.d = "0"
typVBRasEntry.ipAddrWinsAlt.a = "0"
typVBRasEntry.ipAddrWinsAlt.B = "0"
typVBRasEntry.ipAddrWinsAlt.c = "0"
typVBRasEntry.ipAddrWinsAlt.d = "0"
typVBRasEntry.LocalPhoneNumber = "1260"
 
Dim rtn As Long
rtn = VBRasSetEntryProperties("KTUNESCO", typVBRasEntry)
End Sub

Private Function VBRasSetEntryProperties(strEntryName As String, _
         typRasEntry As VBRasEntry, _
         Optional strPhoneBook As String) As Long
   
   Dim rtn As Long, lngCb As Long, lngBuffLen As Long
   Dim B() As Byte
   Dim lngPos As Long, lngStrLen As Long
   
   rtn = RasGetEntryProperties(vbNullString, vbNullString, _
                           ByVal 0&, lngCb, ByVal 0&, ByVal 0&)

   If rtn <> 603 Then VBRasSetEntryProperties = rtn: Exit Function
   
   lngStrLen = Len(typRasEntry.AlternateNumbers)
   lngBuffLen = lngCb + lngStrLen + 1
   ReDim B(lngBuffLen)
   
   CopyMemory B(0), lngCb, 4
   CopyMemory B(4), typRasEntry.options, 4
   CopyMemory B(8), typRasEntry.CountryID, 4
   CopyMemory B(12), typRasEntry.CountryCode, 4
   CopyStringToByte B(16), typRasEntry.AreaCode, 11
   CopyStringToByte B(27), typRasEntry.LocalPhoneNumber, 129
   
   If lngStrLen > 0 Then
     CopyMemory B(lngCb), _
               ByVal typRasEntry.AlternateNumbers, lngStrLen
     CopyMemory B(156), lngCb, 4
   End If

   CopyMemory B(160), typRasEntry.ipAddr, 4
   CopyMemory B(164), typRasEntry.ipAddrDns, 4
   CopyMemory B(168), typRasEntry.ipAddrDnsAlt, 4
   CopyMemory B(172), typRasEntry.ipAddrWins, 4
   CopyMemory B(176), typRasEntry.ipAddrWinsAlt, 4
   CopyMemory B(180), typRasEntry.FrameSize, 4
   CopyMemory B(184), typRasEntry.fNetProtocols, 4
   CopyMemory B(188), typRasEntry.FramingProtocol, 4
   CopyStringToByte B(192), typRasEntry.ScriptName, 260
   CopyStringToByte B(452), typRasEntry.AutodialDll, 260
   CopyStringToByte B(712), typRasEntry.AutodialFunc, 260
   CopyStringToByte B(972), typRasEntry.DeviceType, 17
      If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
   CopyStringToByte B(989), typRasEntry.DeviceName, lngStrLen
      lngPos = 989 + lngStrLen
   CopyStringToByte B(lngPos), typRasEntry.X25PadType, 33
      lngPos = lngPos + 33
   CopyStringToByte B(lngPos), typRasEntry.X25Address, 201
      lngPos = lngPos + 201
   CopyStringToByte B(lngPos), typRasEntry.X25Facilities, 201
      lngPos = lngPos + 201
   CopyStringToByte B(lngPos), typRasEntry.X25UserData, 201
      lngPos = lngPos + 203
   CopyMemory B(lngPos), typRasEntry.Channels, 4
   
   If lngCb > 1768 Then
      CopyMemory B(1768), typRasEntry.NT4En_SubEntries, 4
      CopyMemory B(1772), typRasEntry.NT4En_DialMode, 4
      CopyMemory B(1776), typRasEntry.NT4En_DialExtraPercent, 4
      CopyMemory B(1780), typRasEntry.NT4En_DialExtraSampleSeconds, 4
      CopyMemory B(1784), typRasEntry.NT4En_HangUpExtraPercent, 4
      CopyMemory B(1788), typRasEntry.NT4En_HangUpExtraSampleSeconds, 4
      CopyMemory B(1792), typRasEntry.NT4En_IdleDisconnectSeconds, 4
      
      If lngCb > 1796 Then
         CopyMemory B(1796), typRasEntry.Win2000_Type, 4
         CopyMemory B(1800), typRasEntry.Win2000_EncryptionType, 4
         CopyMemory B(1804), typRasEntry.Win2000_CustomAuthKey, 4
         CopyMemory B(1808), typRasEntry.Win2000_guidId(0), 16
         CopyStringToByte B(1824), typRasEntry.Win2000_CustomDialDll, 260
         CopyMemory B(2084), typRasEntry.Win2000_VpnStrategy, 4
      End If
      
   End If
   
   rtn = RasSetEntryProperties(strPhoneBook, strEntryName, _
                              B(0), lngCb, ByVal 0&, ByVal 0&)
   
   VBRasSetEntryProperties = rtn

End Function

Private Sub CopyStringToByte(bPos As Byte, _
                        strToCopy As String, lngMaxLen As Long)
   Dim lngLen As Long
   lngLen = Len(strToCopy)
   If lngLen = 0 Then
      Exit Sub
   ElseIf lngLen > lngMaxLen Then
      lngLen = lngMaxLen
   End If
   CopyMemory bPos, ByVal strToCopy, lngLen
End Sub

Private Function GetDeviceName() As String
    Dim lpRasDevInfo As RASDEVINFO
    Dim lpcb As Long
    Dim cDevices As Long
    Dim t_Buff As Long
    Dim nRet As Long
    Dim t_ptr As Long
    Dim i As Long
    
    lpcb = 0

    lpRasDevInfo.dwSize = LenB(lpRasDevInfo) + (LenB(lpRasDevInfo) Mod 4)

    nRet = iRasEnumDevices(ByVal 0, lpcb, cDevices)

    t_Buff = GlobalAlloc(GPTR, lpcb)

    iCopyMemory ByVal t_Buff, lpRasDevInfo, LenB(lpRasDevInfo)
 
    nRet = iRasEnumDevices(ByVal t_Buff, lpcb, lpcb)

    If nRet = 0 Then
        t_ptr = t_Buff

        For i = 0 To cDevices - 1
            iCopyMemory lpRasDevInfo, ByVal t_ptr, LenB(lpRasDevInfo)
            If lpRasDevInfo.szDeviceType(0) = Asc("m") And lpRasDevInfo.szDeviceType(1) = Asc("o") And lpRasDevInfo.szDeviceType(2) = Asc("d") And lpRasDevInfo.szDeviceType(3) = Asc("e") And lpRasDevInfo.szDeviceType(4) = Asc("m") Then
                    GetDeviceName = lpRasDevInfo.szDeviceName
                    GoTo KT
            End If
            t_ptr = t_ptr + LenB(lpRasDevInfo) + (LenB(lpRasDevInfo) Mod 4)
        Next i
    Else
        MsgBox "Kh«ng t×m thÊy modem !", vbCritical, App.ProductName
    End If
KT:
    If t_Buff <> 0 Then GlobalFree (t_Buff)
End Function

Public Sub EMailDB()
    Dim fname As String, email As String, smtp As String
    
    email = frmMain.LbCty(9).Caption
    smtp = SelectSQL("SELECT SMTP AS F1 FROM License")
    If InStr(email, "@") = 0 Or Left(smtp, 1) = "." Then
        email = "" '"unesco@hn.vnn.vn"
        smtp = "mail.hn.vnn.vn"
    End If
    DelTemp
    fname = pCurDir + "DB" + CStr(Int_StrToCode(pTenCty)) + CStr(Int_StrToCode(pTenCn)) + "_" + CStr(Day(Date)) + CStr(Month(Date)) + CStr(Year(Date)) + ".SAS"
    Recycle pCurDir + "TEMP.MDB"
    Recycle fname
    CloseUp 1
    DoEvents
    On Error Resume Next
    DBEngine.CompactDatabase pDataPath, pCurDir + "TEMP.MDB", , , ";pwd=" + pPSW
    If Len(Dir(pCurDir + "TEMP.MDB")) > 0 Then
        NenTep pCurDir + "TEMP.MDB", fname
        Recycle pCurDir + "TEMP.MDB"
    Else
        NenTep pDataPath, fname
    End If
    OpenDB pDataPath
    frmMail.SendMail email, smtp, fname
End Sub

'   MIME class
Public Function StartWinSock()
  
  Dim RequestedVersion As Integer
  
  CRLF = Chr(13) + Chr(10)
  'frmVBWinSockDialog.Show
  
  'Debug.Print "StartWinSock"
  RequestedVersion = &H101    ' Version 1.1
  Status = WSAStartup(RequestedVersion, WSAdata)
  
  If Status Then
    'DebugLastError "WinSock startup "
    StartWinSock = False
    Exit Function
  End If
    
  'Debug.Print "Startup status " & SocketError(Status)
  'Debug.Print "Version " & WSAdata.wVersion
  'Debug.Print "High version " & WSAdata.wHighVersion
  'Debug.Print "Description " & WSAdata.szDescription
  'Debug.Print "System status " & WSAdata.szSystemStatus
  'Debug.Print "Max sockets " & WSAdata.iMaxSockets
  'Debug.Print "Max datagrams " & WSAdata.iMaxUdpDg
  
  'Debug.Print "Pointer to vendor info " & WSAdata.lpVendorInfo
  StartWinSock = True

End Function

Function GetIPAddress(HostAddress As Long, HostName As String) As Boolean
    
  Dim nheHostRecord As NameHostEntry
  Dim lngHostRecordPointer As Long
  Dim lngHostAddressPointer As Long
  Dim lngHostAddressPointerPointer As Long
  
  lngHostAddressPointerPointer = 0
  lngHostAddressPointer = 0
  lngHostRecordPointer = 0

  lngHostRecordPointer = gethostbyname(HostName)
  If lngHostRecordPointer <> 0 Then
      hmemcpy nheHostRecord, ByVal lngHostRecordPointer, Len(nheHostRecord)
  End If
  lngHostAddressPointerPointer = nheHostRecord.AddressPointerPointer
  
  If lngHostAddressPointerPointer <> 0 Then
      hmemcpy lngHostAddressPointer, ByVal lngHostAddressPointerPointer&, 4
  End If
  If lngHostAddressPointerPointer <> 0 Then
      hmemcpy HostAddress, ByVal lngHostAddressPointer, 4
  End If
  
  'Debug.Print HostAddress
  dota& = inet_ntoa&(HostAddress)
  dotaddr$ = Space(256)
  temp& = lstrcpy&(dotaddr$, dota&)
  'Debug.Print dotaddr$
  
  GetIPAddress = True
End Function

Function CreateSocket(TCPSocket As Long, ByVal TCPPort As Long)

  'Debug.Print "Creating socket"
  
    'Get a socket/handle from Winsock
  TCPSocket = socket(PF_INET, SOCK_STREAM, AF_UNSPEC)
  If TCPSocket = INVALID_SOCKET Then
    'DebugLastError "Creation of socket"
    CreateSocket = False
    Exit Function
  End If

  'Debug.Print "Got socket handle " & TCPSocket
  
    'Now bind the socket
  'Debug.Print "Bind socket " & TCPSocket
  
    'We pass this structure during a bind.
  CallSocket.Family = AF_INET
  CallSocket.Port = htons(TCPPort)
  CallSocket.Address = 0
  CallSocket.Fill = " "
  
  Status = bind(TCPSocket, CallSocket, Len(CallSocket))
  If Status = SOCKET_ERROR Then
    'DebugLastError "Bind of socket"
    CreateSocket = False
    Exit Function
  End If
  
  'Debug.Print "Bind OK"
  CreateSocket = True
  
End Function

Function ConnectSocket(TCPSocket As Long, HostAddress As Long, ByVal TCPPort As Long)

  CallSocket.Family = AF_INET
  CallSocket.Port = htons(TCPPort)     'converts byte order from PC to Network
  CallSocket.Address = HostAddress
  CallSocket.Fill = " "

  'Debug.Print "Connect socket " & TCPSocket
  Status = connect(TCPSocket, CallSocket, Len(CallSocket))
  If Status = SOCKET_ERROR Then
    'DebugLastError "Connecting socket"
    ConnectSocket = False
    Exit Function
  End If
      
  'Debug.Print "Connected to server OK."
  ConnectSocket = True
  
End Function

Function SendSocketBinary(TCPSocket As Long, FullBuffer As String)

  Dim RemainingBuffer As String
  Dim RemainingLength As Long
  
  DialogPrintBinary "[" + Trim$(str$(TCPSocket)) + "-->]", FullBuffer
  
  RemainingBuffer = FullBuffer
  
  Do
    RemainingLength = Len(RemainingBuffer)
    Status = send(TCPSocket, RemainingBuffer, RemainingLength, 0)
    
    If Status = SOCKET_ERROR Then
      'DebugLastError "Send to socket " + Str(TCPSocket)
      SendSocketBinary = False
      Exit Function
    ElseIf Status = RemainingLength Then
      'Debug.Print "Send was OK."
      SendSocketBinary = True
      Exit Do
    Else
      RemainingBuffer = Mid(RemainingBuffer, Status + 1)
      'Debug.Print "Partial send of " & Status & " bytes"
    End If
  Loop

End Function


Function SendSocket(TCPSocket As Long, FullBuffer As String, Optional OmitCRLF As Boolean) As Boolean
  Dim RemainingBuffer As String
  Dim RemainingLength As Long
  
  'DialogPrint "[" + Trim$(Str$(TCPSocket)) + "-->]", FullBuffer
  'Debug.Print FullBuffer
  
  If OmitCRLF Then
    RemainingBuffer = FullBuffer
  Else
    RemainingBuffer = FullBuffer + Chr$(13) + Chr$(10)
  End If
  
  Do
    RemainingLength = Len(RemainingBuffer)
    Status = send(TCPSocket, RemainingBuffer, RemainingLength, 0)
    
    If Status = SOCKET_ERROR Then
      'DebugLastError "Send to socket " + Str(TCPSocket)
      SendSocket = False
      Exit Function
    ElseIf Status = RemainingLength Then
      'Debug.Print "Send was OK."
      SendSocket = True
      Exit Do
    Else
      RemainingBuffer = Mid(RemainingBuffer, Status + 1)
      'Debug.Print "Partial send of " & Status & " bytes"
    End If
  Loop
  
End Function

Function ReceiveSocketBinary(TCPSocket As Long, SocketResponse As String) As Long
  Dim ResponseLength As Integer
  
      'Now wait for the response from the server.
  SocketResponse = Space$(4096)
  Status = recv(TCPSocket, SocketResponse, 4096, 0)
  
  If Status = SOCKET_ERROR Then
    'DebugLastError "Receive from socket " + Str(TCPSocket)
    ReceiveSocketBinary = False
    Exit Function
  End If
    
  ResponseLength = Status
  SocketResponse = Left(SocketResponse, ResponseLength)
  'Debug.Print "Successful receive of " & ResponseLength & " bytes"
  DialogPrintBinary "[-->" + Trim$(str$(TCPSocket)) + "] ", SocketResponse
  ReceiveSocketBinary = True
  'Debug.Print ReceiveSocketBinary
  
End Function

Function ReceiveSocket(TCPSocket As Long, SocketResponse As String, Optional Port As Integer) As Boolean
  Static strResponseRemainder(10) As Variant
  Dim intLoopCount As Integer
  
  If IsMissing(Port) Then Port = 1
  
  If Port >= 0 Then
    intLoopCount = 0
    Do While InStr(1, strResponseRemainder(Port), CRLF) = 0
      SocketResponse = Space$(4096)
      Do
        Status = recv(TCPSocket, SocketResponse, 4096, 0)
        intLoopCount = intLoopCount + 1
      Loop Until Status <> SOCKET_ERROR Or _
                 WSAGetLastError() <> 0 Or _
                 intLoopCount > 30
                 
      If TCPSocket = 0 Then
        'MsgBox hello
        End
      End If
      
      If (Status = SOCKET_ERROR) Then
        SocketResponse = ""
        'DebugLastError "Receive from socket " + Str(TCPSocket)
        ReceiveSocket = False
        Exit Function
      End If
      'Debug.Print "Successful receive of " & Status & " bytes"
      
      strResponseRemainder(Port) = strResponseRemainder(Port) + Trim(SocketResponse)
    
    Loop
    
    intCRLFPosition = InStr(1, strResponseRemainder(Port), CRLF)
    
    SocketResponse = Left(strResponseRemainder(Port), intCRLFPosition - 1)
    strResponseRemainder(Port) = Mid(strResponseRemainder(Port), intCRLFPosition + 2)
  Else
  
    SocketResponse = Space$(32000)
    Status = recv(TCPSocket, SocketResponse, 32000, 0)
    SocketResponse = Trim(SocketResponse)
    
    If Status = SOCKET_ERROR Then
      'DebugLastError "Receive from socket " + Str(TCPSocket)
      ReceiveSocket = False
      Exit Function
    End If
    'Debug.Print "Successful receive of " & Status & " bytes"
  
  End If
  
  'DialogPrint "[-->" + Trim$(Str$(TCPSocket)) + "] " + Chr$(13) + Chr$(10), SocketResponse
  'Debug.Print SocketResponse
  
  ReceiveSocket = True

End Function

Function ReleaseSocket(TCPSocket As Long)

  'If TCPSocket <= 0 Then
    'Debug.Print "Socket not in use"
  'End If
  
  Status = closesocket(TCPSocket)
  
  If Status = SOCKET_ERROR Then
    'DebugLastError "Close of socket " & Str(TCPSocket)
    ReleaseSocket = False
    Exit Function
  End If
  
  'Debug.Print "Closed socket " & TCPSocket & " OK"
  ReleaseSocket = True

End Function

Sub DialogPrintBinary(ByVal PrintPrefix As String, ByVal PrintText As String)
  Dim PrintLine As String
  Dim X As Integer
  Dim DialogText As String
  
  DialogText = frmVBWinSockDialog!txtDialog.Text
  
  DialogText = DialogText + PrintPrefix + Chr(13) + Chr(10)
  
  For X = 1 To Len(PrintText)
    DialogText = DialogText + Right(Space(4) + str(Asc(Mid$(PrintText, X, 1))), 4)
    If Int(X / 16) = X / 16 Then DialogText = DialogText + Chr(13) + Chr(10)
  Next X
  DialogText = DialogText + Chr(13) + Chr(10) + Chr(13) + Chr(10)
  If Len(DialogText) > 30000 Then DialogText = Right(DialogText, 30000)
    
  frmVBWinSockDialog!txtDialog.Text = DialogText
  frmVBWinSockDialog!txtDialog.Refresh


End Sub

'Sub DialogPrint(ByVal PrintPrefix As String, ByVal PrintText As String)
'    Dim DialogText As String
'    Dim PrintLine As String
'
'    DialogText = frmVBWinSockDialog!txtDialog.Text
'
'    PrintLine = PrintPrefix + " >" + Trim$(PrintText) + "<"
'
'    If Right$(PrintLine, 2) <> Chr$(13) + Chr$(10) Then
'        PrintLine = PrintLine + Chr$(13) + Chr$(10)
'    End If
'
'    DialogText = DialogText + PrintLine
'    If Len(DialogText) > 30000 Then DialogText = Right(DialogText, 30000)
'
'    frmVBWinSockDialog!txtDialog.Text = DialogText
'    frmVBWinSockDialog!txtDialog.Refresh
'
'End Sub

Function SocketError(ErrorCode) As String
  
  Select Case ErrorCode
    Case 0:     SocketError = "OK"
    Case 10004: SocketError = "WSAEINTR"
    Case 10009: SocketError = "WSAEBADF"
    Case 10013: SocketError = "WSAEACCES"
    Case 10014: SocketError = "WSAEFAULT"
    Case 10022: SocketError = "WSAEINVAL"
    Case 10024: SocketError = "WSAEMFILE"
    Case 10035: SocketError = "WSAEWOULDBLOCK"
    Case 10036: SocketError = "WSAEINPROGRESS"
    Case 10037: SocketError = "WSAEALREADY"
    Case 10038: SocketError = "WSAENOTSOCK"
    Case 10039: SocketError = "WSAEDESTADDRREQ"
    Case 10040: SocketError = "WSAEMSGSIZE"
    Case 10041: SocketError = "WSAEPROTOTYPE"
    Case 10042: SocketError = "WSAENOPROTOOPT"
    Case 10043: SocketError = "WSAEPROTONOSUPPORT"
    Case 10044: SocketError = "WSAESOCKTNOSUPPORT"
    Case 10045: SocketError = "WSAEOPNOTSUPP"
    Case 10046: SocketError = "WSAEPFNOSUPPORT"
    Case 10047: SocketError = "WSAEAFNOSUPPORT"
    Case 10048: SocketError = "WSAEADDRINUSE"
    Case 10049: SocketError = "WSAEADDRNOTAVAIL"
    Case 10050: SocketError = "WSAENETDOWN"
    Case 10051: SocketError = "WSAENETUNREACH"
    Case 10052: SocketError = "WSAENETRESET"
    Case 10053: SocketError = "WSAECONNABORTED"
    Case 10054: SocketError = "WSAECONNRESET"
    Case 10055: SocketError = "WSAENOBUFS"
    Case 10056: SocketError = "WSAEISCONN"
    Case 10057: SocketError = "WSAENOTCONN"
    Case 10058: SocketError = "WSAESHUTDOWN"
    Case 10059: SocketError = "WSAETOOMANYREFS"
    Case 10060: SocketError = "WSAETIMEDOUT"
    Case 10061: SocketError = "WSAECONNREFUSED"
    Case 10062: SocketError = "WSAELOOP"
    Case 10063: SocketError = "WSAENAMETOOLONG"
    Case 10064: SocketError = "WSAEHOSTDOWN"
    Case 10065: SocketError = "WSAEHOSTUNREACH"
    Case 10066: SocketError = "WSAENOTEMPTY"
    Case 10067: SocketError = "WSAEPROCLIM"
    Case 10068: SocketError = "WSAEUSERS"
    Case 10069: SocketError = "WSAEDQUOT"
    Case 10070: SocketError = "WSAESTALE"
    Case 10071: SocketError = "WSAEREMOTE"
    Case 10091: SocketError = "WSASYSNOTREADY"
    Case 10092: SocketError = "WSAVERNOTSUPPORTED"
    Case 10093: SocketError = "WSANOTINITIALIZED"
    Case 11001: SocketError = "WSAEWSA_HOST_NOT_FOUND"
    Case 11002: SocketError = "WSATRY_AGAIN"
    Case 11003: SocketError = "WSANO_RECOVERY"
    Case 11004: SocketError = "WSANO_DATA"
    Case Else
       SocketError = "UNKNOWN ERROR CODE" + str$(ErrorCode)
  End Select
End Function

Function LastSocketError() As String
  Status = WSAGetLastError()
  LastSocketError = SocketError(Status)
End Function



