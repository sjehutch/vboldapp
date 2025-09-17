Attribute VB_Name = "modGlobals"
Option Explicit


'constants that need to be declared in the Globals file
Public Const REG_KEY_PATH = "Software\DAS\Profume\Registration"
Public Const SETTING_VAL_COMPANY = "Company"
Public Const SETTING_VAL_ADDRESS1 = "Address1"
Public Const SETTING_VAL_ADDRESS2 = "Address2"
Public Const SETTING_VAL_ADDRESS3 = "Address3"
Public Const SETTING_VAL_CITY = "City"
Public Const SETTING_VAL_STATE = "State"
Public Const SETTING_VAL_ZIP = "ZIP"
Public Const SETTING_VAL_COUNTRY = "Country"
Public Const SETTING_VAL_NAME = "Name"

'***********************************************************************
Public Const GRID_SETTINGS_PATH  As String = "Software\DAS\ProFume\Settings"
Public Const GRID_MAX_ROWS As Integer = 101     'number of rows in the grid. This includes the header row
Public Const RTF_FONT_SIZE As Integer = 8      'the font size of the Results string
Public Const RTF_FONT_NAME As String = "arial"

Public Const MIN_HUMIDITY As Double = 1
Public Const MAX_HUMIDITY As Double = 90
' 9/25/2003 SIMS #2103035 Jett changed minimum line length to get more
' accurate results (25 ft/8 meters)

'Pat 1/4/2005 Added and Removed constants
'Public Const MIN_LINE_LENGTH_METRIC As Double = 8
'Public Const MIN_LINE_LENGTH_ENGLISH As Double = 25
'Public Const MAX_LINE_LENGTH_METRIC As Double = 152
'Public Const MAX_LINE_LENGTH_ENGLISH As Double = 500

Public Const MIN_LENGTH_ONEFOURTH_METRIC As Double = 0
Public Const MIN_LENGTH_ONEFOURTH_ENGLISH As Double = 0
Public Const MAX_LENGTH_ONEFOURTH_METRIC As Double = 213.36
Public Const MAX_LENGTH_ONEFOURTH_ENGLISH As Double = 700
Public Const MIN_LENGTH_ONEEIGHT_METRIC As Double = 0
Public Const MIN_LENGTH_ONEEIGHT_ENGLISH As Double = 0
Public Const MAX_LENGTH_ONEEIGHT_METRIC As Double = 30.48
Public Const MAX_LENGTH_ONEEIGHT_ENGLISH As Double = 100

Public Const MIN_PRES_METRIC As Double = 0
Public Const MAX_PRES_METRIC As Double = 30
Public Const MIN_PRES_ENGLISH As Double = 0
Public Const MAX_PRES_ENGLISH As Double = 446
'*********


Public Const MIN_FAN_METRIC As Double = 28.31
Public Const MIN_FAN_ENGLISH As Double = 1000
Public Const MAX_FAN_METRIC As Double = 707.92
Public Const MAX_FAN_ENGLISH As Double = 25000
Public Const POUNDS_PER_KG As Double = 2.2046226
Public Const MIN_LOAD_FACTOR As Integer = 0
Public Const MIN_TEMP_METRIC As Double = 4.44
Public Const MIN_TEMP_ENGLISH As Double = 40
Public Const MAX_TEMP_METRIC As Double = 65.56
Public Const MAX_TEMP_ENGLISH As Double = 150
Public Const MIN_EST_HLT As Double = 1#
Public Const MAX_EST_HLT As Double = 1000#
Public Const MIN_PLANNED_EXPOSURE As Double = 1#
Public Const MAX_PLANNED_EXPOSURE As Double = 168
Public Const MIN_VOLUME_METRIC As Double = 1#
Public Const MIN_VOLUME_ENGLISH  As Double = 35
Public Const MAX_VOLUME_METRIC As Double = 283168.47
Public Const MAX_VOLUME_ENGLISH As Double = 10000000
Public Const MAX_LOAD_FACTOR As Integer = 100
Public Const DEFAULT_TEMPERATURE As Double = 23.89
Public Const DEFAULT_HLT As Double = MIN_EST_HLT
Public Const DEFAULT_EXPOSURE As Double = MIN_PLANNED_EXPOSURE
Public Const DEFAULT_VOLUME As Double = 1#

Public Const GRID_COLOR = "&H80000008"
Public Const BACK_COLOR_ALTERNATE = "&H80000016"
Public Const READ_ONLY_COLOR = "&H8000000B"
Public Const COLUMN_HEADING_COLOR = "&H808080"
'Public Const HIGHLIGHT_COLOR = "&HFFFF00"
Public Const HIGHLIGHT_COLOR = BACK_COLOR_ALTERNATE
Public Const COL_INDEX_WIDTH As Integer = 400
'**************************************************************************8

'--Monitor Status Constants
Public Const STATUS_UNKNOWN = 0
Public Const STATUS_ON_TARGET = 1
Public Const STATUS_ABOVE_TARGET = 2
Public Const STATUS_BELOW_TARGET = 3
Public Const STATUS_ACHIEVED_TARGET = 4
Public Const STATUS_INCREASING_CONCENTRATION = 5
Public Const STATUS_TARGET_DOSE_INCOMPATABLE = 6
Public Const STATUS_BAD_ELAPSED_TIMES = 7
Public Const STATUS_MISSING_CONC_READING = 8
Public Const DBL_MAX = 1.79769313486232E+307  ' Maximum Value
Public Const MDM_UNKNOWN = 0
Public Const MDM_PRE_INTRO = 1
Public Const MDM_END_INTRO = 2
Public Const MDM_PRE_EXPOSURE = 3
Public Const MDM_EXPOSURE = 4
Public Const MDM_PRE_AERATION = 5
Public Const MDM_AERATION = 6
Public Const MDM_POST_CLEARANCE = 7
Public Const MDM_MARK_EXPOSURE = 14
Public Const MDM_MARK_PRE_AERATION = 15
Public Const MDM_MARK_AERATION = 16
Public Const MDM_MARK_POST_CLEARANCE = 17

Public Const CRYPKEY_DAYS_TILL_EXPIRE = 30 ' Crypkey - days till expiration

' copied from the Main module
Public gsCurrentFile As String
Public gsTempFile As String
Public gbHasChanged As Boolean
Public gbFileReadOnly As Boolean
Public gbIsFileOpen As Boolean
Public gdbConn As ADODB.Connection
Public gbCanChangeTabs As Boolean
Public gbIsTestingActive As Boolean
Public gForceLanguage As clLanguages ' JETT 2/11/04 LOCALIZATION - Added this variable to track forced language
Public gTimeFormat As String   ' JETT 3/5/04 LOCALIZATION - Added this variable to track time format
Public gFullTimeFormat As String  ' JETT 3/5/04 LOCALIZATION - Added this variable to track time format
Public gbCanExit As Boolean
Public gbIsExiting As Boolean
Public gfrmActiveForm As Form
Public gbIsMetric As Boolean
Public gsJobID As String
'--MME Changed gnExposureTime data type to DOUBLE (Used to be integer)
Public gnExposureTime As Double     ' Exposure time for this job

'New Global variables for the GRAPH
Public gbShowLines As Boolean
Public gnGraphUnits As Integer

'JNM 05/02/03 - PT fix - define new global to tell the app if it is saving or not
Public gbIsSaving As Boolean



' ---------- added by jett Fumiguide+ Project

Public gActiveFrame As Frame

Public blnTempPres As Boolean
Public gblPresTempMode As Boolean
Public gblPressure As Boolean
Public blnEnable As Boolean
Public blnConvert As Boolean
Public blnChangeTab As Boolean

'PAT 2/10/2006 - Added
Public gbChkCommodity As Boolean
Public gbChkLoadFactor As Boolean
Public gbChkEnabled As Boolean


