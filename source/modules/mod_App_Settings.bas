Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Settings
' Level:        Application module
' Version:      1.11
' Description:  Application-wide related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015  - 1.00 - initial version
'               BLC, 11/20/2015 - 1.01 - added priority & status icons
'               BLC, 6/7/2016   - 1.02 - updated documentation & added ACCESS_ROLES (Big Rivers App)
'               BLC, 6/20/2016  - 1.03 - added DB_ADMIN_FORM and documentation
'               BLC, 9/1/2016   - 1.04 - updated APP_SYS_TABLES
'               BLC, 9/7/2016   - 1.05 - added LINK_NORMAL_TEXT for disabling links
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added updated version to Upland db
' --------------------------------------------------------------------
'               BLC, 3/22/2017  - 1.06 - removed big rivers only components
'                                        revised for uplands
' --------------------------------------------------------------------
'               BLC, 4/24/2017          added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 5/3/2017  - 1.08 - added VCS_FULL_PATH for running VCS functions/subroutines
'               BLC, 7/5/2017  - 1.09 - added QUADRATS_PER_TRANSECT to make
'                                       adding quadrats for new transects flexible in case
'                                       # changes from 3 quadrats per transect
'               BLC, 7/12/2017 - 1.10 - added VCS_SAVE_TABLES for tables to backup (lookups)
'               BLC, 7/28/2017 - 1.11 - changed DEV_MODE to global variable vs. constant to
'                                       allow user to set via tglDevMode toggle control in UI
' =================================

' ---------------------------------
' GLOBALS:      global values set for application
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, June 2016
' Adapted:      -
' Revisions:    BLC, 6/6/2016 - initial version (NCPN WQ Big Rivers App, App_Templates)
' ---------------------------------
'Public g_AppTemplates As Scripting.Dictionary     'global dictionary for application templates (if any)

Public gSubReportCount As Integer                  'global counter for subreports

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, May 2014
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version (NCPN WQ Utilities Tool, WATER_YEAR_START & WATER_YEAR_END)
'               BLC, 4/22/2015 - adapted to generic tools (NCPN Invasives Reporting Tool) by adding
'                                USER_ACCESS_CONTROL (False - gives users full control in apps w/o controls,
'                                                     True - relies on user access control settings)
'                                DB_SYS_TABLES & APP_SYS_TABLES (handle table arrays for the database/
'                                   application)
'               BLC, 4/30/2015 - add DB_ADMIN_CONTROL flag to handle applications w/o full DbAdmin subform & controls
'                                add MAIN_APP_FORM constant to handle applications where frm_Switchboard is NOT the main form
'                                add APP_RELEASE_ID constant to handle application release ID w/o full DbAdmin subfrom & controls
'               BLC, 5/1/2015  - add DEV_MODE constant to enable menus typically off during use
'               BLC, 5/13/2015 - shifted UI enable/disabled colors from TempVars set in initialize (mod_App_UI) to constants
'               BLC, 5/19/2015 - added FIX_LINKED_DBS flag to handle applications which require updates of tbl_Dbs via FixLinkedDb
'                                (usually when DbAdmin is not fully implemented)
'               BLC, 5/28/2015 - added MAIN_APP_MENU to handle applications w/ main menu forms (not tabbed switchboards)
'               BLC, 4/4/2016  - added LOCATION_TYPES to allow specific types only, RECORD_ACTIONS, CONTACT_ROLES, PARKS
'               BLC, 6/7/2016  - added ACCESS_ROLES to set user application permissions
'               BLC, 9/7/2016  - added LINK_NORMAL_TEXT & _BKGD for disabling tile links
'               BLC, 5/3/2017  - added VCS_FULL_PATH for running VCS functions/subroutines
'               BLC, 7/28/2017 - changed DEV_MODE to global variable vs. constant to
'                                allow user to set via tglDevMode toggle control in UI
' ---------------------------------
Public Const VCS_FULL_PATH As String = "Z:\_____LIB\dev\git_projects\libraries\VCS.accdb"   'Version Control System (VCS) db (contains modules for version control)
                                                                'Tables to save for VCS (e.g. lookups)
Public Const VCS_SAVE_TABLES As String = "tlu_projects, " & _
    "tlu_Cover_Code, tlu_Cover_Value, tlu_Disturbance," & _
    "tlu_NCPN_Plants, tlu_Parks, tlu_Size_Value, tlu_Species"
Public Const USER_ACCESS_CONTROL As Boolean = True              'Boolean flag -> db includes user access control or not
Public Const DB_ADMIN_CONTROL As Boolean = False                'Boolean flag -> db does not include DbAdmin subform & controls
Public Const FIX_LINKED_DBS As Boolean = False                  'Boolean flag -> db requires tbl_Dbs to be updated via FixLinkedDb (usually when DbAdmin is not fully implemented)
Public Const MAIN_APP_FORM As String = "Main"                   'String -> main tabbed form (frm_Switchboard, etc.)
Public Const MAIN_APP_MENU As String = "Main"                   'String -> main tabbed form (frm_Switchboard, etc.)
Public Const APP_RELEASE_ID As String = ""                      'String -> release ID (tsys_App_Release.Release_ID) for current release
                                                                '          used when db doesn't include full DbAdmin subform & controls, otherwise NULL
Public Const APP_URL As String = "science.nature.nps.gov/im/units/ncpn/datamanagement.cfm"
                                                                'String -> website URL for application
                                                                '          used when db doesn't include full DbAdmin subform & controls, otherwise NULL
Public DEV_MODE As Boolean                                      'Boolean flag -> enable menus, show controls when typically they'd be OFF
                                                                '        flag is set via DEV_MODE toggle in UI

Public Const ACCESS_ROLES As String = "admin,power user,data entry,read only"
                                                                'String -> used in setting user application access level & permissions
Public Const SWITCHBOARD As String = "Main"                     'String -> main application form
Public Const DB_ADMIN_FORM As String = "DbAdmin"                'String -> main db administrative form
Public Const BACKEND_REQUIRED As Boolean = True                 'Boolean flag -> identifies if back-end required

'-----------------------------------------------------------------------
' Database Type
'-----------------------------------------------------------------------
Public Const BACKEND_TYPE As String = "ACCESS"

'-----------------------------------------------------------------------
' Database System Tables
'-----------------------------------------------------------------------
'   Array("App_Defaults", "BE_Updates", "Link_Dbs", "Link_Tables")
'   tsys_App_Defaults -> default application settings
'   tsys_BE_Updates   -> updates to post to remot back-end copies
'   tsys_Link_Dbs     -> info about linked back-end dbs
'   tsys_Link_Tables  -> info about linked tables
'-----------------------------------------------------------------------
' Application Backend System Tables
'-----------------------------------------------------------------------
'   Array("App_Releases", "Bug_Reports", "Logins", "User_Roles")
'   tsys_App_Releases -> list of application releases
'   tsys_Bug_Reports  -> tracking for known issues
'   tsys_Logins       -> system use monitoring
'   tsys_User_Roles   -> assign user access priviledges  [deprecated to Contact_Access]
'-----------------------------------------------------------------------
' SEE ALSO >>>> SysTablesExist() function
'-----------------------------------------------------------------------
Public Const DB_SYS_TABLES As String = "App_Defaults, Link_Files, Link_Tables"
Public Const APP_SYS_TABLES As String = "App_Releases, Bug_Reports, Logins"

'-----------------------------------------------------------------------
' User Interface Colors
'-----------------------------------------------------------------------
'std control colors
Public Const CTRL_DISABLED As Long = lngLtGray
Public Const CTRL_ADD_ENABLED As Long = lngLime
Public Const CTRL_REMOVE_ENABLED As Long = lngLtOrange
Public Const TEXT_ENABLED As Long = lngBlue
Public Const TEXT_DISABLED As Long = lngGray

'highlight text for tile links
Public Const LINK_HIGHLIGHT_TEXT As Long = lngBlue
Public Const LINK_HIGHLIGHT_BKGD As Long = lngYelLime
Public Const HIGHLIGHT_MISSING_VALUE As Long = lngYellow
Public Const LINK_NORMAL_TEXT As Long = lngGray50

Public Const PROGRESS_BAR As Long = lngLime

'-----------------------------------------------------------------------
' Icons
'-----------------------------------------------------------------------
Public Const ICON_PATH As String = "Z:\_____LIB\dev\git_projects\icons\small\"

Public Const FLAG_RED As String = ICON_PATH & "flag_red" & ".png"
Public Const FLAG_LIME As String = ICON_PATH & "flag_lime" & ".png"
Public Const FLAG_ORANGE As String = ICON_PATH & "flag_orange" & ".png"
Public Const FLAG_LTBLUE As String = ICON_PATH & "flag_ltblue" & ".png"
Public Const FLAG_BLUE As String = ICON_PATH & "flag_blue" & ".png"
Public Const FLAG_NAVY As String = ICON_PATH & "flag_navy" & ".png"
Public Const FLAG_PURPLE As String = ICON_PATH & "flag_purple" & ".png"

Public Const DOT_RED As String = ICON_PATH & "dot_red" & ".png"
Public Const DOT_LIME As String = ICON_PATH & "dot_lime" & ".png"
Public Const DOT_ORANGE As String = ICON_PATH & "dot_orange" & ".png"
Public Const DOT_LTBLUE As String = ICON_PATH & "dot_ltblue" & ".png"
Public Const DOT_BLUE As String = ICON_PATH & "dot_blue" & ".png"
Public Const DOT_NAVY As String = ICON_PATH & "dot_navy" & ".png"
Public Const DOT_PURPLE As String = ICON_PATH & "dot_purple" & ".png"

'-----------------------------------------------------------------------
' Photo Types
'-----------------------------------------------------------------------
Public Const PHOTO_TYPES_MAIN As String = "Reference,Overview,Feature,Transect,Other"      'String -> basic photo types
Public Const PHOTO_TYPES_OTHER As String = "Animal,Plant,Cultural,Disturbance,Field Work,Scenic,Weather,Other"      'String -> other photo types
Public Const PHOTO_EXT_ALLOWED As String = "jpg,jpeg,png"
Public Const PHOTO_PATH As String = "Z:\_____LIB\dev\git_projects\big_rivers_app\"
'photo number regex pattern defined in AppEnum

'-----------------------------------------------------------------------
' Upland Components
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' Invasives Components
'-----------------------------------------------------------------------
Public g_AppSurfaces As Scripting.Dictionary            'global application surface names & IDs (for lookups)
Public g_AppQuadratPositions As Scripting.Dictionary    'global application quadrat positions (for lookups)

Public Const QUADRATS_PER_TRANSECT As Integer = 3       'total # of quadrats found on an invasives transect
                                                        'this value assumes quadrat #s are consecutive & begin w/ 1