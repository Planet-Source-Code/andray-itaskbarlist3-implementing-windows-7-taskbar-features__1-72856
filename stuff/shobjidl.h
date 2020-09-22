/*
 *  shobjidl.h  Shell object interfaces
 *
:include crwatcnt.sp
 */

#include <rpc.h>
#include <rpcndr.h>
#ifndef COM_NO_WINDOWS_H
    #include <windows.h>
    #include <ole2.h>
#endif

#ifndef __shobjidl_h__
#define __shobjidl_h__

:include readonly.sp

#include <objidl.h>
#include <oleidl.h>
#include <oaidl.h>
#include <docobj.h>
#include <shtypes.h>
#include <servprov.h>
#include <comcat.h>
#include <propidl.h>
#include <prsht.h>
#include <msxml.h>
#include <propsys.h>
#include <sherrors.h>
#include <commctrl.h>
#ifndef RC_INVOKED
    #include <objectarray.h>
#endif

:include cpluspro.sp

/* Forward declarations */
typedef interface IShellView            IShellView;
typedef interface IShellBrowser         IShellBrowser;
typedef interface IShellFolder          IShellFolder;
typedef interface IShellItem            IShellItem;
#if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
typedef interface IShellItemArray       IShellItemArray;
#endif
#if (NTDDI_VERSION >= 0x06000000)
typedef interface IFileDialog           IFileDialog;
typedef interface IShellItemFilter      IShellItemFilter;
#endif
typedef interface IEnumExplorerCommand  IEnumExplorerCommand;
typedef interface ITransferAdviseSink   ITransferAdviseSink;
typedef interface IEnumShellItems       IEnumShellItems;
typedef interface IShellItemFilter      IShellItemFilter;
typedef interface IContextMenuCB        IContextMenuCB;

/* Shell primitive data types */
typedef ULONG                   SFGAOF;
typedef char                    *LPVIEWSETTINGS;
typedef LPFNADDPROPSHEETPAGE    LPFNSVADDPROPSHEETPAGE;
typedef GUID                    SHELLVIEWID;
typedef HRESULT                 DEPRECATED_HRESULT;
typedef LPTBBUTTON              LPTBBUTTONSB;
#if (NTDDI_VERSION >= 0x05010000)
typedef GUID                    STGTRANSCONFIRMATION;
typedef GUID                    *LPSTGTRANSCONFIRMATION;
#endif
#if (NTDDI_VERSION >= 0x06000000)
typedef GUID                    EXPLORERPANE;
#endif
typedef HANDLE                  HTHEME;

/* Reference data types */
#if (NTDDI_VERSION >= 0x06000000)
    #ifdef __cplusplus
        #define REFEXPLORERPANE const EXPLORERPANE &
    #else
        #define REFEXPLORERPANE const EXPLORERPANE * const
    #endif
#endif

/* Context menu flags */
#define CMF_NORMAL              0x00000000
#define CMF_DEFAULTONLY         0x00000001
#define CMF_VERBSONLY           0x00000002
#define CMF_EXPLORE             0x00000004
#define CMF_NOVERBS             0x00000008
#define CMF_CANRENAME           0x00000010
#define CMF_NODEFAULT           0x00000020
#if (NTDDI_VERSION < 0x06000000)
    #define CMF_INCLUDESTATIC   0x00000040
#endif
#if (NTDDI_VERSION >= 0x06000000)
    #define CMF_ITEMMENU        0x00000080
#endif
#define CMF_EXTENDEDVERBS       0x00000100
#if (NTDDI_VERSION >= 0x06000000)
    #define CMF_DISABLEDVERBS   0x00000200
#endif
#define CMF_ASYNCVERBSTATE      0x00000400
#define CMF_OPTIMIZEFORINVOKE   0x00000800
#define CMF_SYNCCASCADEMENU     0x00001000
#define CMF_DONOTPICKDEFAULT    0x00002000

/* IContextMenu::GetCommandString() flags */
#define GCS_VERBA           0x00000000
#define GCS_HELPTEXTA       0x00000001
#define GCS_VALIDATEA       0x00000002
#define GCS_VERBW           0x00000004
#define GCS_HELPTEXTW       0x00000005
#define GCS_VALIDATEW       0x00000006
#define GCS_VERBICONW       0x00000014
#define GCS_UNICODE         0x00000004
#ifdef UNICODE
    #define GCS_VERB        GCS_VERBW
    #define GCS_HELPTEXT    GCS_HELPTEXTW
    #define GCS_VALIDATE    GCS_VALIDATEW
#else
    #define GCS_VERB        GCS_VERBA
    #define GCS_HELPTEXT    GCS_HELPTEXTA
    #define GCS_VALIDATE    GCS_VALIDATEA
#endif

/* Command strings */
#define CMDSTR_NEWFOLDERA       "NewFolder"
#define CMDSTR_VIEWLISTA        "ViewList"
#define CMDSTR_VIEWDETAILSA     "ViewDetails"
#define CMDSTR_NEWFOLDERW       L"NewFolder"
#define CMDSTR_VIEWLISTW        L"ViewList"
#define CMDSTR_VIEWDETAILSW     L"ViewDetails"
#ifdef UNICODE
    #define CMDSTR_NEWFOLDER    CMDSTR_NEWFOLDERW
    #define CMDSTR_VIEWLIST     CMDSTR_VIEWLISTW
    #define CMDSTR_VIEWDETAILS  CMDSTR_VIEWDETAILSW
#else
    #define CMDSTR_NEWFOLDER    CMDSTR_NEWFOLDERA
    #define CMDSTR_VIEWLIST     CMDSTR_VIEWLISTA
    #define CMDSTR_VIEWDETAILS  CMDSTR_VIEWDETAILSA
#endif

/* IContextMenu::InvokeCommand() masks */
#define CMIC_MASK_HOTKEY                SEE_MASK_HOTKEY
#define CMIC_MASK_ICON                  SEE_MASK_ICON
#define CMIC_MASK_FLAG_NO_UI            SEE_MASK_FLAG_NO_UI
#define CMIC_MASK_UNICODE               SEE_MASK_UNICODE
#define CMIC_MASK_NO_CONSOLE            SEE_MASK_NO_CONSOLE
#if (NTDDI_VERSION < 0x06000000)
    #define CMIC_MASK_HASLINKNAME       SEE_MASK_HASLINKNAME
    #define CMIC_MASK_HASTITLE          SEE_MASK_HASTITLE
#endif
#define CMIC_MASK_FLAG_SEP_VDM          SEE_MASK_FLAG_SEPVDM
#define CMIC_MASK_ASYNCOK               SEE_MASK_ASYNCOK
#if (NTDDI_VERSION >= 0x06000000)
    #define CMIC_MASK_NOASYNC           SEE_MASK_NOASYNC
#endif
#if (_WIN32_IE >= 0x0501)
    #define CMIC_MASK_SHIFT_DOWN        0x10000000L
#endif
#if (_WIN32_IE >= 0x0400)
    #define CMIC_MASK_PTINVOKE          0x20000000L
#endif
#if (_WIN32_IE >= 0x0501)
    #define CMIC_MASK_CONTROL_DOWN      0x40000000L
#endif
#if (_WIN32_IE >= 0x0560)
    #define CMIC_MASK_FLAG_LOG_USAGE    SEE_MASK_FLAG_LOG_USAGE
    #define CMIC_MASK_NOZONECHECKS      SEE_MASK_NOZONECHECKS
#endif

/* IRunnableTask states */
#if (_WIN32_IE >= 0x0400)
    #define IRTIR_TASK_NOT_RUNNING  0
    #define IRTIR_TASK_RUNNING      1
    #define IRTIR_TASK_SUSPENDED    2
    #define IRTIR_TASK_PENDING      3
    #define IRTIR_TASK_FINISHED     4
#endif

/* IShellTaskScheduler special values */
#if (_WIN32_IE >= 0x0400)
    #define TOID_NULL               GUID_NULL
    #define ITSAT_DEFAULT_LPARAM    ((DWORD_PTR)-1)
#endif

/* IShellTaskScheduler::AddTask() priorities */
#if (_WIN32_IE >= 0x0400)
    #define ITSAT_DEFAULT_PRIORITY  0x10000000L
    #define ITSAT_MAX_PRIORITY      0x7FFFFFFFL
    #define ITSAT_MIN_PRIORITY      0x00000000L
#endif

/* IShellTaskScheduler::Status() flags */
#if (_WIN32_IE >= 0x0400)
    #define ITSSFLAG_COMPLETE_ON_DESTROY    0x0000
    #define ITSSFLAG_KILL_ON_DESTROY        0x0001
    #define ITSSFLAG_FLAGS_MASK             0x0003
#endif

/* IShellTaskScheduler::Status() timeouts */
#if (_WIN32_IE >= 0x0400)
    #define ITSS_THREAD_DESTROY_DEFAULT_TIMEOUT 10000
    #define ITSS_THREAD_TERMINATE_TIMEOUT       INFINITE
    #define ITSS_THREAD_TIMEOUT_NO_CHANGE       (INFINITE - 1)
#endif

/* IShellFolder::CompareIDs() flags */
#define SHCIDS_ALLFIELDS        0x80000000L
#define SHCIDS_CANONICALONLY    0x10000000L
#define SHCIDS_BITMASK          0xFFFF0000L
#define SHCIDS_COLUMNMASK       0x0000FFFFL

/* IShellFolder::GetAttributesOf() flags */
#define SFGAO_CANCOPY           DROPEFFECT_COPY
#define SFGAO_CANMOVE           DROPEFFECT_MOVE
#define SFGAO_CANLINK           DROPEFFECT_LINK
#define SFGAO_STORAGE           0x00000008L
#define SFGAO_CANRENAME         0x00000010L
#define SFGAO_CANDELETE         0x00000020L
#define SFGAO_HASPROPSHEET      0x00000040L
#define SFGAO_DROPTARGET        0x00000100L
#define SFGAO_CAPABILITYMASK    0x00000177L
#define SFGAO_SYSTEM            0x00001000L
#define SFGAO_ENCRYPTED         0x00002000L
#define SFGAO_ISSLOW            0x00004000L
#define SFGAO_GHOSTED           0x00008000L
#define SFGAO_LINK              0x00010000L
#define SFGAO_SHARE             0x00020000L
#define SFGAO_READONLY          0x00040000L
#define SFGAO_HIDDEN            0x00080000L
#define SFGAO_DISPLAYATTRMASK   0x000FC000L
#define SFGAO_NONENUMERATED     0x00100000L
#define SFGAO_NEWCONTENT        0x00200000L
#define SFGAO_CANMONIKER        0x00400000L
#define SFGAO_HASSTORAGE        0x00400000L
#define SFGAO_STREAM            0x00400000L
#define SFGAO_STORAGEANCESTOR   0x00800000L
#define SFGAO_VALIDATE          0x01000000L
#define SFGAO_REMOVABLE         0x02000000L
#define SFGAO_COMPRESSED        0x04000000L
#define SFGAO_BROWSABLE         0x08000000L
#define SFGAO_FILESYSANCESTOR   0x10000000L
#define SFGAO_FOLDER            0x20000000L
#define SFGAO_FILESYSTEM        0x40000000L
#define SFGAO_STORAGECAPMASK    0x70C50008L
#define SFGAO_HASSUBFOLDER      0x80000000L
#define SFGAO_CONTENTSMASK      0x80000000L
#define SFGAO_PKEYSFGAOMASK     0x81044000L

/* IShellFolder bind strings */
#define STR_BIND_FORCE_FOLDER_SHORTCUT_RESOLVE  L"Force Folder Shortcut Resolve"
#define STR_AVOID_DRIVE_RESTRICTION_POLICY      L"Avoid Drive Restriction Policy"
#define STR_SKIP_BINDING_CLSID                  L"Skip Binding CLSID"
#define STR_PARSE_PREFER_FOLDER_BROWSING        L"Parse Prefer Folder Browsing"
#define STR_DONT_PARSE_RELATIVE                 L"Don't Parse Relative"
#define STR_PARSE_TRANSLATE_ALIASES             L"Parse Translate Aliases"
#define STR_PARSE_SKIP_NET_CACHE                L"Skip Net Resource Cache"
#define STR_PARSE_SHELL_PROTOCOL_TO_FILE_OBJECTS \
    L"Parse Shell Protocol To File Objects"
#if (_WIN32_IE >= 0x0700)
    #define STR_TRACK_CLSID                     L"Track the CLSID"
    #define STR_INTERNAL_NAVIGATE               L"Internal Navigation"
    #define STR_PARSE_PROPERTYSTORE             L"DelegateNamedProperties"
    #define STR_NO_VALIDATE_FILENAME_CHARS      L"NoValidateFilenameChars"
    #define STR_IFILTER_LOAD_DEFINED_FILTER \
        L"Only bind registered persistent handlers"
    #define STR_IFILTER_FORCE_TEXT_FILTER_FALLBACK \
        L"Always bind persistent handlers"
    #define STR_BIND_DELEGATE_CREATE_OBJECT     L"Delegate Object Creation"
    #define STR_PARSE_ALLOW_INTERNET_SHELL_FOLDERS \
        L"Allow binding to Internet shell folder handlers and negate STR_PARSE_PREFER_WEB_BROWSING"
    #define STR_PARSE_PREFER_WEB_BROWSING \
        L"Do not bind to Internet shell folder handlers"
    #define STR_PARSE_SHOW_NET_DIAGNOSTICS_UI   L"Show network diagnostics UI"
    #define STR_PARSE_DONT_REQUIRE_VALIDATED_URLS \
        L"Do not require validated URLs"
    #define STR_INTERNETFOLDER_PARSE_ONLY_URLMON_BINDABLE   L"Validate URL"
#endif
#if (NTDDI_VERSION >= 0x06010000)
    #define STR_BIND_FOLDERS_READ_ONLY          L"Folders As Read Only"
    #define STR_BIND_FOLDER_ENUM_MODE           L"Folder Enum Mode"
    #define STR_PARSE_WITH_EXPLICIT_PROGID      L"ExplicitProgid"
    #define STR_PARSE_WITH_EXPLICIT_ASSOCAPP    L"ExplicitAssociationApp"
    #define STR_PARSE_EXPLICIT_ASSOCIATION_SUCCESSFUL   L"ExplicitAssociationSuccessful"
    #define STR_PARSE_AND_CREATE_ITEM           L"ParseAndCreateItem"
    #define STR_ITEM_CACHE_CONTEXT              L"ItemCacheContext"
#endif
#define STR_FILE_SYS_BIND_DATA                  L"File System Bind Data"

/* IShellView2::GetView() special values */
#define SV2GV_CURRENTVIEW   ((UINT)-1)
#define SV2GV_DEFAULTVIEW   ((UINT)-1)

/* ICommDlgBrowser::OnStateChange() values */
#define CDBOSC_SETFOCUS     0x00000000L
#define CDBOSC_KILLFOCUS    0x00000001L
#define CDBOSC_SELCHANGE    0x00000002L
#define CDBOSC_RENAME       0x00000003L
#define CDBOSC_STATECHANGE  0x00000004L

/* ICommDlgBrowser2::Notify() notification types */
#if (NTDDI_VERSION >= 0x05000000)
    #define CDB2N_CONTEXTMENU_DONE  0x00000001L
    #define CDB2N_CONTEXTMENU_START 0x00000002L
#endif

/* ICommDlgBrowser2::GetViewFlags() flags */
#if (NTDDI_VERSION >= 0x05000000)
    #define CDB2GVF_SHOWALLFILES        0x00000001L
#endif
#if (NTDDI_VERSION >= 0x06000000)
    #define CDB2GVF_ISFILESAVE          0x00000002L
    #define CDB2GVF_ALLOWPREVIEWPANE    0x00000004L
    #define CDB2GVF_NOSELECTVERB        0x00000008L
    #define CDB2GVF_NOINCLUDEITEM       0x00000010L
    #define CDB2GVF_ISFOLDERPICKER      0x00000020L
    #define CDB2GVF_ADDSHIELD           0x00000040L
#endif

/* Maximum column name length */
#if (_WIN32_IE >= 0x0700)
    #define MAX_COLUMN_NAME_LEN 80
#endif

/* IShellBrowser::BrowseObject() flags */
#define SBSP_DEFBROWSER                 0x00000000
#define SBSP_SAMEBROWSER                0x00000001
#define SBSP_NEWBROWSER                 0x00000002
#define SBSP_DEFMODE                    0x00000000
#define SBSP_OPENMODE                   0x00000010
#define SBSP_EXPLOREMODE                0x00000020
#define SBSP_HELPMODE                   0x00000040
#define SBSP_NOTRANSFERHIST             0x00000080
#define SBSP_ABSOLUTE                   0x00000000
#define SBSP_RELATIVE                   0x00001000
#define SBSP_PARENT                     0x00002000
#define SBSP_NAVIGATEBACK               0x00004000
#define SBSP_NAVIGATEFORWARD            0x00008000
#define SBSP_ALLOW_AUTONAVIGATE         0x00010000
#if (NTDDI_VERSION >= 0x06000000)
    #define SBSP_KEEPSAMETEMPLATE       0x00020000
    #define SBSP_KEEPWORDWHEELTEXT      0x00040000
    #define SBSP_ACTIVATE_NOFOCUS       0x00080000
    #define SBSP_CREATENOHISTORY        0x00100000
    #define SBSP_PLAYNOSOUND            0x00200000
#endif
#if (_WIN32_IE >= 0x0602)
    #define SBSP_CALLERUNTRUSTED        0x00800000
    #define SBSP_TRUSTFIRSTDOWNLOAD     0x01000000
    #define SBSP_UNTRSUTEDFORDOWNLOAD   0x02000000
#endif
#define SBSP_NOAUTOSELECT               0x04000000
#define SBSP_WRITENOHISTORY             0x08000000
#if (_WIN32_IE >= 0x0602)
    #define SBSP_TRUSTEDFORACTIVEX      0x10000000
#endif
#if (_WIN32_IE >= 0x0700)
    #define SBSP_FEEDNAVIGATION         0x20000000
#endif
#define SBSP_REDIRECT                   0x40000000
#define SBSP_INITIATEDBYHLINKFRAME      0x80000000

/* IShellBrowser window identifiers */
#define FCW_STATUS      0x0001
#define FCW_TOOLBAR     0x0002
#define FCW_TREE        0x0003
#define FCW_INTERNETBAR 0x0006
#define FCW_PROGRESS    0x0008

/* IShellBrowser toolbar item flags */
#define FCT_MERGE       0x0001
#define FCT_CONFIGABLE  0x0002
#define FCT_ADDTOEND    0x0004

/* IShellItem strings */
#define STR_DONT_RESOLVE_LINK   L"Don't Resolve Link"
#define STR_GET_ASYNC_HANDLER   L"GetAsyncHandler"

/* Property store strings */
#define STR_GPS_HANDLERPROPERTIESONLY   L"GPS_HANDLERPROPERTIESONLY"
#define STR_GPS_FASTPROPERTIESONLY      L"GPS_FASTPROPERTIESONLY"
#define STR_GPS_OPENSLOWITEM            L"GPS_OPENSLOWITEM"
#define STR_GPS_DELAYCREATION           L"GPS_DELAYCREATION"
#define STR_GPS_BESTEFFORT              L"GPS_BESTEFFORT"
#define STR_GPS_NO_OPLOCK               L"GPS_NO_OPLOCK"

/* Shell drag image registered message */
#if (NTDDI_VERSION >= 0x05000000)
    #define DI_GETDRAGIMAGE TEXT( "ShellGetDragImage" )
#endif

/* AutoRun content flags */
#define ARCONTENT_AUTORUNINF            0x00000002L
#define ARCONTENT_AUDIOCD               0x00000004L
#define ARCONTENT_DVDMOVIE              0x00000008L
#define ARCONTENT_BLANKCD               0x00000010L
#define ARCONTENT_BLANKDVD              0x00000020L
#define ARCONTENT_UNKNOWNCONTENT        0x00000040L
#define ARCONTENT_AUTOPLAYPIX           0x00000080L
#define ARCONTENT_AUTOPLAYMUSIC         0x00000100L
#define ARCONTENT_AUTOPLAYVIDEO         0x00000200L
#if (NTDDI_VERSION >= 0x06000000)
    #define ARCONTENT_VCD               0x00000400L
    #define ARCONTENT_SVCD              0x00000800L
    #define ARCONTENT_DVDAUDIO          0x00001000L
    #define ARCONTENT_BLANKBD           0x00002000L
    #define ARCONTENT_BLURAY            0x00004000L
    #define ARCONTENT_NONE              0x00008000L
    #define ARCONTENT_MASK              0x00007FFEL
    #define ARCONTENT_PHASE_UNKNOWN     0x00000000L
    #define ARCONTENT_PHASE_PRESNIFF    0x10000000L
    #define ARCONTENT_PHASE_SNIFFING    0x20000000L
    #define ARCONTENT_PHASE_FINAL       0x40000000L
    #define ARCONTENT_PHASE_MASK        0x70000000L
#endif

/* IExtractImage priorities */
#if (_WIN32_IE >= 0x0400)
    #define IEI_PRIORITY_MAX        ITSAT_MAX_PRIORITY
    #define IEI_PRIORITY_MIN        ITSAT_MIN_PRIORITY
    #define IEIT_PRIORITY_NORMAL    ITSAT_DEFAULT_PRIORITY
#endif

/* IExtractImage flags */
#if (_WIN32_IE >= 0x0400)
    #define IEIFLAG_ASYNC       0x00000001L
    #define IEIFLAG_CACHE       0x00000002L
    #define IEIFLAG_ASPECT      0x00000004L
    #define IEIFLAG_OFFLINE     0x00000008L
    #define IEIFLAG_GLEAM       0x00000010L
    #define IEIFLAG_SCREEN      0x00000020L
    #define IEIFLAG_ORIGSIZE    0x00000040L
    #define IEIFLAG_NOSTAMP     0x00000080L
    #define IEIFLAG_NOBORDER    0x00000100L
    #define IEIFLAG_QUALITY     0x00000200L
    #define IEIFLAG_REFRESH     0x00000400L
#endif

/* Desk band information masks */
#define DBIM_MINSIZE    0x00000001L
#define DBIM_MAXSIZE    0x00000002L
#define DBIM_INTEGRAL   0x00000004L
#define DBIM_ACTUAL     0x00000008L
#define DBIM_TITLE      0x00000010L
#define DBIM_MODEFLAGS  0x00000020L
#define DBIM_BKCOLOR    0x00000040L

/* Desk band information mode flags */
#define DBIMF_NORMAL            0x00000000L
#define DBIMF_FIXED             0x00000001L
#define DBIMF_FIXEDBMP          0x00000004L
#define DBIMF_VARIABLEHEIGHT    0x00000008L
#define DBIMF_UNDELETEABLE      0x00000010L
#define DBIMF_DEBOSSED          0x00000020L
#define DBIMF_BKCOLOR           0x00000040L
#define DBIMF_USECHEVRON        0x00000080L
#define DBIMF_BREAK             0x00000100L
#define DBIMF_ADDTOFRONT        0x00000200L
#define DBIMF_TOPALIGN          0x00000400L
#if (NTDDI_VERSION >= 0x06000000)
    #define DBIMF_NOGRIPPER     0x00000800L
    #define DBIMF_ALWAYSGRIPPER 0x00001000L
    #define DBIMF_NOMARGINS     0x00002000L
#endif

/* IDeskBand::GetBandInfo() flags */
#define DBIF_VIEWMODE_NORMAL        0x0000
#define DBIF_VIEWMODE_VERTICAL      0x0001
#define DBIF_VIEWMODE_FLOATING      0x0002
#define DBIF_VIEWMODE_TRANSPARENT   0x0004

/* Other desk band constants */
#define DBPC_SELECTFIRST    0xFFFFFFFFL
#define DBPC_SELECTLAST     0xFFFFFFFEL

/* Thumb button notifications */
#define THBN_CLICKED    0x1800

/* Wizard extension dialog identifiers */
#define IDD_WIZEXTN_FIRST   0x5000
#define IDD_WIZEXTN_LAST    0x5100

/* Shell publishing wizard host flags */
#define SHPWHF_NORECOMPRESS             0x00000001
#define SHPWHF_NONETPLACECREATE         0x00000002
#define SHPWHF_NOFILESELECTOR           0x00000004
#define SHPWHF_USEMRU                   0x00000008
#if (NTDDI_VERSION >= 0x06000000)
    #define SHPWHF_ANYLOCATION          0x00000100
#endif
#define SHPWHF_VALIDATEVIAWEBFOLDERS    0x00010000

/* Extended file operation flags */
#if (_WIN32_IE >= 0x0700)
    #define FOFX_NOSKIPJUNCTIONS        0x00010000L
    #define FOFX_PREFERHARDLINK         0x00020000L
    #define FOFX_SHOWELEVATIONPROMPT    0x00040000L
    #define FOFX_EARLYFAILURE           0x00100000L
    #define FOFX_PRESERVEFILEEXTENSIONS 0x00200000L
    #define FOFX_KEEPNEWERFILE          0x00400000L
    #define FOFX_NOCOPYHOOKS            0x00800000L
    #define FOFX_NOMINIMIZEBOX          0x01000000L
    #define FOFX_MOVEACLSACROSSVOLUMES  0x02000000L
    #define FOFX_DONTDISPLAYSOURCEPATH  0x04000000L
    #define FOFX_DONTDISPLAYDESTPATH    0x08000000L
    #define FOFX_REQUIREELEVATION       0x10000000L
    #define FOFX_COPYASDOWNLOAD         0x20000000L
    #define FOFX_DONTDISPLAYLOCATIONS   0x40000000L
#endif

/* IAutoCompleteDropDown::GetDropDownStatus() flags */
#define ACDD_VISIBLE    0x00000001L

/* Band site information masks */
#if (_WIN32_IE >= 0x0400)
    #define BSIM_STATE  0x00000001L
    #define BSIM_STYLE  0x00000002L
#endif

/* Band site state flags */
#if (_WIN32_IE >= 0x0400)
    #define BSSF_VISIBLE        0x00000001L
    #define BSSF_NOTITLE        0x00000002L
    #define BSSF_UNDELETEABLE   0x00001000L
#endif

/* Band site item styles */
#if (_WIN32_IE >= 0x0400)
    #define BSIS_AUTOGRIPPER                0x00000000L
    #define BSIS_NOGRIPPER                  0x00000001L
    #define BSIS_ALWAYSGRIPPER              0x00000002L
    #define BSIS_LEFTALIGN                  0x00000004L
    #define BSIS_SINGLECLICK                0x00000008L
    #define BSIS_NOCONTEXTMENU              0x00000010L
    #define BSIS_NODROPTARGET               0x00000020L
    #define BSIS_NOCAPTION                  0x00000040L
    #define BSIS_PREFERNOLINEBREAK          0x00000080L
    #define BSIS_LOCKED                     0x00000100L
#endif
#if (_WIN32_IE >= 0x0700)
    #define BSIS_PRESERVEORDERDURINGLAYOUT  0x00000200L
    #define BSIS_FIXEDORDER                 0x00000400L
#endif

/* CD burning extension property string */
#if (NTDDI_VERSION >= 0x05010000)
    #define PROPSTR_EXTENSIONCOMPLETIONSTATE    L"ExtensionCompletionState"
#endif

/* Open file capabilities */
#if (NTDDI_VERSION >= 0x06000000)
    #define OF_CAP_CANSWITCHTO  0x00000001L
    #define OF_CAP_CANCLOSE     0x00000002L
#endif

/* Shell menu data masks */
#if (_WIN32_IE >= 0x0600)
    #define SMDM_SHELLFOLDER    0x00000001L
    #define SMDM_HMENU          0x00000002L
    #define SMDM_TOOLBAR        0x00000004L
#endif

/* Shell menu callback actions */
#if (_WIN32_IE >= 0x0600)
    #define SMC_INITMENU                0x00000001
    #define SMC_CREATE                  0x00000002
    #define SMC_EXITMENU                0x00000003
    #define SMC_GETINFO                 0x00000005
    #define SMC_GETSFINFO               0x00000006
    #define SMC_GETOBJECT               0x00000007
    #define SMC_GETSFOBJECT             0x00000008
    #define SMC_SFEXEC                  0x00000009
    #define SMC_SFSELECTITEM            0x0000000A
    #define SMC_REFRESH                 0x00000010
    #define SMC_DEMOTE                  0x00000011
    #define SMC_PROMOTE                 0x00000012
    #define SMC_DEFAULTICON             0x00000016
    #define SMC_NEWITEM                 0x00000017
    #define SMC_CHEVRONEXPAND           0x00000019
    #define SMC_DISPLAYCHEVRONTIP       0x0000002A
    #define SMC_SETSFOBJECT             0x0000002D
    #define SMC_SHCHANGENOTIFY          0x0000002E
    #define SMC_CHEVRONGETTIP           0x0000002F
    #define SMC_SFDDRESTRICTED          0x00000030
#endif
#if (_WIN32_IE >= 0x0700)
    #define SMC_SFEXEC_MIDDLE           0x00000031
    #define SMC_GETAUTOEXPANDSTATE      0x00000041
    #define SMC_AUTOEXPANDCHANGE        0x00000042
    #define SMC_GETCONTEXTMENUMODIFIER  0x00000043
    #define SMC_GETBKCONTEXTMENU        0x00000044
    #define SMC_OPEN                    0x00000045
#endif

/* SMC_GETAUTOEXPANDSTATE return values */
#if (_WIN32_IE >= 0x0700)
    #define SMAE_EXPANDED   0x00000001
    #define SMAE_CONTRACTED 0x00000002
    #define SMAE_USER       0x00000004
    #define SMAE_VALID      0x00000007
#endif

/* IShellMenu::Initialize() flags */
#if (_WIN32_IE >= 0x0600)
    #define SMINIT_DEFAULT              0x00000000L
    #define SMINIT_REGISTER_DRAGDROP    0x00000002L
    #define SMINIT_TOPLEVEL             0x00000004L
    #define SMINIT_CACHED               0x00000010L
#endif
#if (_WIN32_IE >= 0x0700)
    #define SMINIT_AUTOEXPAND           0x00000100L
    #define SMINIT_AUTOTOOLTIP          0x00000200L
    #define SMINIT_DROPONCONTAINER      0x00000400L
#endif
#if (_WIN32_IE >= 0x0600)
    #define SMINIT_VERTICAL             0x10000000L
    #define SMINIT_HORIZONTAL           0x20000000L
#endif

/* Ancestor special value */
#if (_WIN32_IE >= 0x0600)
    #define ANCESTORDEFAULT ((UINT)-1)
#endif

/* IShellMenu::SetMenu() flags */
#if (_WIN32_IE >= 0x0600)
    #define SMSET_DONTOWN   0x00000001
    #define SMSET_TOP       0x10000000
    #define SMSET_BOTTOM    0x20000000
#endif

/* IShellMenu::InvalidateItem() flags */
#if (_WIN32_IE >= 0x0600)
    #define SMINV_REFRESH   0x00000001
    #define SMINV_ID        0x00000008
#endif

/* INameSpaceTreeControlDropHandler position special value */
#define NSTCDHPOS_ONTOP (-1)

/* IPreviewHandler::DoPreview() error codes */
#define E_PREVIEWHANDLER_DRM_FAIL   0x86420001L
#define E_PREVIEWHANDLER_NOAUTH     0x86420002L
#define E_PREVIEWHANDLER_NOTFOUND   0x86420003L
#define E_PREVIEWHANDLER_CORRUPT    0x86420004L

/* HomeGroup security group */
#define HOMEGROUP_SECURITY_GROUP    L"HomeUsers"

/* IContextMenu::InvokeCommand() parameters */
#include <pshpack8.h>
typedef struct _CMINVOKECOMMANDINFO {
    DWORD   cbSize;
    DWORD   fMask;
    HWND    hwnd;
    LPCSTR  lpVerb;
    LPCSTR  lpParameters;
    LPCSTR  lpDirectory;
    int     nShow;
    DWORD   dwHotKey;
    HANDLE  hIcon;
} CMINVOKECOMMANDINFO;
typedef CMINVOKECOMMANDINFO         *LPCMINVOKECOMMANDINFO;
typedef const CMINVOKECOMMANDINFO   *LPCCMINVOKECOMMANDINFO;
#include <poppack.h>

/* IContextMenu::InvokeCommand() parameters (extended version) */
#include <pshpack8.h>
typedef struct _CMINVOKECOMMANDINFOEX {
    DWORD   cbSize;
    DWORD   fMask;
    HWND    hwnd;
    LPCSTR  lpVerb;
    LPCSTR  lpParameters;
    LPCSTR  lpDirectory;
    int     nShow;
    DWORD   dwHotKey;
    HANDLE  hIcon;
    LPCSTR  lpTitle;
    LPCWSTR lpVerbW;
    LPCWSTR lpParametersW;
    LPCWSTR lpDirectoryW;
    LPCWSTR lpTitleW;
    POINT   ptInvoke;
} CMINVOKECOMMANDINFOEX;
typedef CMINVOKECOMMANDINFOEX       *LPCMINVOKECOMMANDINFOEX;
typedef const CMINVOKECOMMANDINFOEX *PCCMINVOKECOMMANDINFOEX;
#include <poppack.h>

/* Persist folder target information */
#if (_WIN32_IE >= 0x0500)
#include <pshpack8.h>
typedef struct _PERSIST_FOLDER_TARGET_INFO {
    PIDLIST_ABSOLUTE    pidlTargetFolder;
    WCHAR               szTargetParsingName[MAX_PATH];
    WCHAR               szNetworkProvider[MAX_PATH];
    DWORD               dwAttributes;
    int                 csidl;
} PERSIST_FOLDER_TARGET_INFO;
#include <poppack.h>
#endif

/* IShellFolder::GetDisplayNameOf() flags */
enum _SHGDNF {
    SHGDN_NORMAL        = 0x0000,
    SHGDN_INFOLDER      = 0x0001,
    SHGDN_FOREDITING    = 0x1000,
    SHGDN_FORADDRESSBAR = 0x4000,
    SHGDN_FORPARSING    = 0x8000
};
typedef DWORD   SHGDNF;

/* IShellFolder::EnumObjects() flags */
enum _SHCONTF {
    SHCONTF_CHECKING_FOR_CHILDREN   = 0x00000010,
    SHCONTF_FOLDERS                 = 0x00000020,
    SHCONTF_NONFOLDERS              = 0x00000040,
    SHCONTF_INCLUDEHIDDEN           = 0x00000080,
    SHCONTF_INIT_ON_FIRST_NEXT      = 0x00000100,
    SHCONTF_NETPRINTERSRCH          = 0x00000200,
    SHCONTF_SHAREABLE               = 0x00000400,
    SHCONTF_STORAGE                 = 0x00000800,
    SHCONTF_NAVIGATION_ENUM         = 0x00001000,
    SHCONTF_FASTITEMS               = 0x00002000,
    SHCONTF_FLATLIST                = 0x00004000,
    SHCONTF_ENABLE_ASYNC            = 0x00008000,
    SHCONTF_INCLUDESUPERHIDDEN      = 0x00010000
};
typedef DWORD   SHCONTF;

/* Folder enumeration modes */
#if (NTDDI_VERSION >= 0x06010000)
typedef enum FOLDER_ENUM_MODE {
    FEM_VIEWRESULT  = 0,
    FEM_NAVIGATION  = 1
} FOLDER_ENUM_MODE;
#endif

/* Extra search */
typedef struct EXTRASEARCH {
    GUID    guidSearch;
    WCHAR   wszFriendlyName[80];
    WCHAR   wszUrl[2084];
} EXTRASEARCH;
typedef EXTRASEARCH *LPEXTRASEARCH;

/* Folder flags */
typedef enum FOLDERFLAGS {
    FWF_NONE                = 0x00000000,
    FWF_AUTOARRANGE         = 0x00000001,
    FWF_ABBREVIATEDNAMES    = 0x00000002,
    FWF_SNAPTOGRID          = 0x00000004,
    FWF_OWNERDATA           = 0x00000008,
    FWF_BESTFITWINDOW       = 0x00000010,
    FWF_DESKTOP             = 0x00000020,
    FWF_SINGLESEL           = 0x00000040,
    FWF_NOSUBFOLDERS        = 0x00000080,
    FWF_TRANSPARENT         = 0x00000100,
    FWF_NOCLIENTEDGE        = 0x00000200,
    FWF_NOSCROLL            = 0x00000400,
    FWF_ALIGNLEFT           = 0x00000800,
    FWF_NOICONS             = 0x00001000,
    FWF_SHOWSELALWAYS       = 0x00002000,
    FWF_NOVISIBLE           = 0x00004000,
    FWF_SINGLECLICKACTIVATE = 0x00008000,
    FWF_NOWEBVIEW           = 0x00010000,
    FWF_HIDEFILENAMES       = 0x00020000,
    FWF_CHECKSELECT         = 0x00040000,
    FWF_NOENUMREFRESH       = 0x00080000,
    FWF_NOGROUPING          = 0x00100000,
    FWF_FULLROWSELECT       = 0x00200000,
    FWF_NOFILTERS           = 0x00400000,
    FWF_NOCOLUMNHEADER      = 0x00800000,
    FWF_NOHEADERINALLVIEWS  = 0x01000000,
    FWF_EXTENDEDTILES       = 0x02000000,
    FWF_TRICHECKSELECT      = 0x04000000,
    FWF_AUTOCHECKSELECT     = 0x08000000,
    FWF_NOBROWSERVIEWSTATE  = 0x10000000,
    FWF_SUBSETGROUPS        = 0x20000000,
    FWF_USESEARCHFOLDER     = 0x40000000,
    FWF_ALLOWRTLREADING     = 0x80000000
} FOLDERFLAGS;

/* Folder view modes */
typedef enum FOLDERVIEWMODE {
    FVM_AUTO        = -1,
    FVM_FIRST       = 1,
    FVM_ICON        = 1,
    FVM_SMALLICON   = 2,
    FVM_LIST        = 3,
    FVM_DETAILS     = 4,
    FVM_THUMBNAIL   = 5,
    FVM_TILE        = 6,
    FVM_THUMBSTRIP  = 7,
    FVM_CONTENT     = 8,
    FVM_LAST        = 8
} FOLDERVIEWMODE;

/* Folder logical view modes */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum FOLDERLOGICALVIEWMODE {
    FLVM_UNSPECIFIED    = -1,
    FLVM_FIRST          = 1,
    FLVM_DETAILS        = 1,
    FLVM_TILES          = 2,
    FLVM_ICONS          = 3,
    FLVM_LIST           = 4,
    FLVM_CONTENT        = 5,
    FLVM_LAST           = 5
} FOLDERLOGICALVIEWMODE;
#endif

/* Folder settings */
typedef struct FOLDERSETTINGS {
    UINT    ViewMode;
    UINT    fFlags;
} FOLDERSETTINGS;
typedef FOLDERSETTINGS          *LPFOLDERSETTINGS;
typedef const FOLDERSETTINGS    *LPCFOLDERSETTINGS;
typedef FOLDERSETTINGS          *PFOLDERSETTINGS;

/* Folder view options */
typedef enum FOLDERVIEWOPTIONS {
    FVO_DEFAULT             = 0x00000000,
    FVO_VISTALAYOUT         = 0x00000001,
    FVO_CUSTOMPOSITION      = 0x00000002,
    FVO_CUSTOMORDERING      = 0x00000004,
    FVO_SUPPORTHYPERLINKS   = 0x00000008,
    FVO_NOANIMATIONS        = 0x00000010,
    FVO_NOSCROLLTIPS        = 0x00000020
} FOLDERVIEWOPTIONS;

/* IShellView::SelectItem() flags */
typedef enum _SVSIF {
    SVSI_DESELECT       = 0x00000000,
    SVSI_SELECT         = 0x00000001,
    SVSI_EDIT           = 0x00000003,
    SVSI_DESELECTOTHERS = 0x00000004,
    SVSI_ENSUREVISIBLE  = 0x00000008,
    SVSI_FOCUSED        = 0x00000010,
    SVSI_TRANSLATEPT    = 0x00000020,
    SVSI_SELECTIONMARK  = 0x00000040,
    SVSI_POSITIONITEM   = 0x00000080,
    SVSI_CHECK          = 0x00000100,
    SVSI_CHECK2         = 0x00000200,
    SVSI_KEYBOARDSELECT = 0x00000401,
    SVSI_NOTAKEFOCUS    = 0x40000000
} _SVSIF;
typedef UINT    SVSIF;
#define SVSI_NOSTATECHANGE  0x80000000

/* IShellView::GetItemObject() flags */
typedef enum _SVGIO {
    SVGIO_BACKGROUND        = 0x00000000,
    SVGIO_SELECTION         = 0x00000001,
    SVGIO_ALLVIEW           = 0x00000002,
    SVGIO_CHECKED           = 0x00000003,
    SVGIO_TYPE_MASK         = 0x0000000F,
    SVGIO_FLAG_VIEWORDER    = 0x80000000
} _SVGIO;
typedef int SVGIO;

/* IShellView::UIActivate() status values */
typedef enum SVUIA_STATUS {
    SVUIA_DEACTIVATE        = 0,
    SVUIA_ACTIVATE_NOFOCUS  = 1,
    SVUIA_ACTIVATE_FOCUS    = 2,
    SVUIA_INPLACEACTIVATE   = 3
} SVUIA_STATUS;

/* IShellView2::CreateViewWindow2() parameters */
#include <pshpack8.h>
typedef struct _SV2CVW2_PARAMS {
    DWORD               cbSize;
    IShellView          *psvPrev;
    LPCFOLDERSETTINGS   pfs;
    IShellBrowser       *psbOwner;
    RECT                *prcView;
    const SHELLVIEWID   *pvid;
    HWND                hwndView;
} SV2CVW2_PARAMS;
typedef SV2CVW2_PARAMS  *LPSV2CVW2_PARAMS;
#include <poppack.h>

/* IShellView3::CreateViewWindow3() flags */
#if (NTDDI_VERSION >= 0x06000000)
enum _SV3CVW3_FLAGS {
    SV3CVW3_DEFAULT             = 0x00000000,
    SV3CVW3_NONINTERACTIVE      = 0x00000001,
    SV3CVW3_FORCEVIEWMODE       = 0x00000002,
    SV3CVW3_FORCEFOLDERFLAGS    = 0x00000004
};
typedef DWORD   SV3CVW3_FLAGS;
#endif

/* Sort directions */
#if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
enum tagSORTDIRECTION {
    SORT_DESCENDING = -1,
    SORT_ASCENDING  = 1
};
typedef int SORTDIRECTION;
#endif

/* Sort column */
#if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
typedef struct SORTCOLUMN {
    PROPERTYKEY     propkey;
    SORTDIRECTION   direction;
} SORTCOLUMN;
#endif

/* Folder view text types */
#if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
typedef enum FVTEXTTYPE {
    FVST_EMPTYTEXT  = 0
} FVTEXTTYPE;
#endif

/* IVisualProperties watermark flags */
#if (_WIN32_IE >= 0x0700)
typedef enum VPWATERMARKFLAGS {
    VPWF_DEFAULT    = 0x00000000,
    VPWF_ALPHABLEND = 0x00000001
} VPWATERMARKFLAGS;
#endif

/* IVisualProperties color flags */
#if (_WIN32_IE >= 0x0700)
typedef enum VPCOLORFLAGS {
    VPCF_TEXT           = 1,
    VPCF_BACKGROUND     = 2,
    VPCF_SORTCOLUMN     = 3,
    VPCF_SUBTEXT        = 4,
    VPCF_TEXTBACKGROUND = 5
} VPCOLORFLAGS;
#endif

/* IColumnManager masks */
#if (_WIN32_IE >= 0x0700)
typedef enum CM_MASK {
    CM_MASK_WIDTH           = 0x00000001,
    CM_MASK_DEFAULTWIDTH    = 0x00000002,
    CM_MASK_IDEALWIDTH      = 0x00000004,
    CM_MASK_NAME            = 0x00000008,
    CM_MASK_STATE           = 0x00000010
} CM_MASK;
#endif

/* IColumnManager states */
#if (_WIN32_IE >= 0x0700)
typedef enum CM_STATE {
    CM_STATE_NONE               = 0x00000000,
    CM_STATE_VISIBLE            = 0x00000001,
    CM_STATE_FIXEDWIDTH         = 0x00000002,
    CM_STATE_NOSORTBYFOLDERNESS = 0x00000004,
    CM_STATE_ALWAYSVISIBLE      = 0x00000008
} CM_STATE;
#endif

/* IColumnManager enumeration flags */
#if (_WIN32_IE >= 0x0700)
typedef enum CM_ENUM_FLAGS {
    CM_ENUM_ALL     = 0x00000001,
    CM_ENUM_VISIBLE = 0x00000002
} CM_ENUM_FLAGS;
#endif

/* IColumnManager set width values */
#if (_WIN32_IE >= 0x0700)
typedef enum CM_SET_WIDTH_VALUE {
    CM_WIDTH_USEDEFAULT = -1,
    CM_WIDTH_AUTOSIZE   = -2
} CM_SET_WIDTH_VALUE;
#endif

/* IColumnManager column information */
#if (_WIN32_IE >= 0x0700)
typedef struct CM_COLUMNINFO {
    DWORD   cbSize;
    DWORD   dwMask;
    DWORD   dwState;
    UINT    uWidth;
    UINT    uDefaultWidth;
    UINT    uIdealWidth;
    WCHAR   wszName[MAX_COLUMN_NAME_LEN];
} CM_COLUMNINFO;
#endif

/* IShellItem::GetDisplayName() flags */
typedef enum _SIGDN {
    SIGDN_NORMALDISPLAY                 = 0x00000000,
    SIGDN_PARENTRELATIVEPARSING         = 0x80018001,
    SIGDN_DESKTOPABSOLUTEPARSING        = 0x80028000,
    SIGDN_PARENTRELATIVEEDITING         = 0x80031001,
    SIGDN_DESKTOPABSOLUTEEDITING        = 0x8004C000,
    SIGDN_FILESYSPATH                   = 0x80058000,
    SIGDN_URL                           = 0x80068000,
    SIGDN_PARENTRELATIVEFORADDRESSBAR   = 0x8007C001,
    SIGDN_PARENTRELATIVE                = 0x80080001
} SIGDN;

/* IShellItem::Compare() hint flags */
enum _SICHINTF {
    SICHINT_DISPLY                          = 0x00000000,
    SICHINT_ALLFIELDS                       = 0x80000000,
    SICHINT_CANONICAL                       = 0x10000000,
    SICHINT_TEST_FILESYSPATH_IF_NOT_EQUAL   = 0x20000000
};
typedef DWORD   SICHINTF;

/* SHGetItemFromDataObject() flags */
#if (NTDDI_VERSION >= 0x06010000)
typedef enum DATAOBJ_GET_ITEM_FLAGS {
    DOGIF_DEFAULT       = 0x00000000,
    DOGIF_TRAVERSE_LINK = 0x00000001,
    DOGIF_NO_HDROP      = 0x00000002,
    DOGIF_NO_URL        = 0x00000004,
    DOGIF_ONLY_IF_ONE   = 0x00000008
} DATAOBJ_GET_ITEM_FLAGS;
#endif

/* IShellItemImageFactory::GetImage() flags */
enum _SIIGB {
    SIIGBF_RESIZETOFIT      = 0x00000000,
    SIIGBF_BIGGERSIZEOK     = 0x00000001,
    SIIGBF_MEMORYONLY       = 0x00000002,
    SIIGBF_ICONONLY         = 0x00000004,
    SIIGBF_THUMBNAILONLY    = 0x00000008,
    SIIGBF_INCACHEONLY      = 0x00000010
};
typedef int SIIGBF;

/* Storage operations */
#if (NTDDI_VERSION >= 0x05010000)
typedef enum STGOP {
    STGOP_MOVE              = 1,
    STGOP_COPY              = 2,
    STGOP_SYNC              = 3,
    STGOP_REMOVE            = 5,
    STGOP_RENAME            = 6,
    STGOP_APPLYPROPERTIES   = 8,
    STGOP_NEW               = 10
} STGOP;
#endif

/* Transfer source flags */
enum _TRANSFER_SOURCE_FLAGS {
    TSF_NORMAL                      = 0x00000000,
    TSF_FAIL_EXIST                  = 0x00000000,
    TSF_RENAME_EXIST                = 0x00000001,
    TSF_OVERWRITE_EXIST             = 0x00000002,
    TSF_ALLOW_DECRYPTION            = 0x00000004,
    TSF_NO_SECURITY                 = 0x00000008,
    TSF_COPY_CREATION_TIME          = 0x00000010,
    TSF_COPY_WRITE_TIME             = 0x00000020,
    TSF_USE_FULL_ACCESS             = 0x00000040,
    TSF_DELETE_RECYCLE_IF_POSSIBLE  = 0x00000080,
    TSF_COPY_HARD_LINK              = 0x00000100,
    TSF_COPY_LOCALIZED_NAME         = 0x00000200,
    TSF_MOVE_AS_COPY_DELETE         = 0x00000400,
    TSF_SUSPEND_SHELLEVENTS         = 0x00000800
};
typedef DWORD   TRANSFER_SOURCE_FLAGS;

/* Transfer advise states */
#if (_WIN32_IE >= 0x0700)
enum _TRANSFER_ADVISE_STATE {
    TS_NONE             = 0x00000000,
    TS_PERFORMING       = 0x00000001,
    TS_PREPARING        = 0x00000002,
    TS_INDETERMINATE    = 0x00000004
};
typedef DWORD   TRANSFER_ADVISE_STATE;
#endif

/* Shell item resource */
typedef struct SHELL_ITEM_RESOURCE {
    GUID    guidType;
    WCHAR   szName[260];
} SHELL_ITEM_RESOURCE;

/* Shell item attribute flags */
typedef enum SIATTRIBFLAGS {
    SIATTRIBFLAGS_AND       = 0x00000001,
    SIATTRIBFLAGS_OR        = 0x00000002,
    SIATTRIBFLAGS_APPCOMPAT = 0x00000003,
    SIATTRIBFLAGS_MASK      = 0x00000003,
    SIATTRIBFLAGS_ALLITEMS  = 0x00004000
} SIATTRIBFLAGS;

/* Property UI name flags */
enum _PROPERTYUI_NAME_FLAGS {
    PUIFNF_DEFAULT  = 0x00000000,
    PUIFNF_MNEMONIC = 0x00000001
};
typedef DWORD   PROPERTYUI_NAME_FLAGS;

/* Property UI flags */
enum _PROPERTYUI_FLAGS {
    PUIF_DEFAULT            = 0x00000000,
    PUIF_RIGHTALIGN         = 0x00000001,
    PUIF_NOLABELININFOTIP   = 0x00000002
};
typedef DWORD   PROPERTYUI_FLAGS;

/* Property UI format flags */
enum _PROPERTYUI_FORMAT_FLAGS {
    PUIFFDF_DEFAULT         = 0x00000000,
    PUIFFDF_RIGHTTOLEFT     = 0x00000001,
    PUIFFDF_SHORTFORMAT     = 0x00000002,
    PUIFFDF_NOTIME          = 0x00000004,
    PUIFFDF_FRIENDLYDATE    = 0x00000008
};
typedef DWORD   PROPERTYUI_FORMAT_FLAGS;

/* Category information flags */
#if (_WIN32_IE >= 0x0500)
typedef enum CATEGORYINFO_FLAGS {
    CATINFO_NORMAL          = 0x00000000,
    CATINFO_COLLAPSED       = 0x00000001,
    CATINFO_HIDDEN          = 0x00000002,
    CATINFO_EXPANDED        = 0x00000004,
    CATINFO_NOHEADER        = 0x00000008,
    CATINFO_NOTCOLLAPSIBLE  = 0x00000010,
    CATINFO_NOHEADERCOUNT   = 0x00000020,
    CATINFO_SUBSETTED       = 0x00000040
} CATEGORYINFO_FLAGS;
#endif

/* Category sort flags */
#if (_WIN32_IE >= 0x0500)
typedef enum CATSORT_FLAGS {
    CATSORT_DEFAULT = 0x00000000,
    CATSORT_NAME    = 0x00000001
} CATSORT_FLAGS;
#endif

/* Category information */
#if (_WIN32_IE >= 0x0500)
typedef struct CATEGORY_INFO {
    CATEGORYINFO_FLAGS  cif;
    WCHAR               wszName[260];
} CATEGORY_INFO;
#endif

/* Shell drag image */
#if (NTDDI_VERSION >= 0x05000000)
#include <pshpack8.h>
typedef struct SHDRAGIMAGE {
    SIZE        sizeDragImage;
    POINT       ptOffset;
    HBITMAP     hbmpDragImage;
    COLORREF    crColorKey;
} SHDRAGIMAGE;
typedef SHDRAGIMAGE *LPSHDRAGIMAGE;
#include <poppack.h>
#endif

/* IDragSourceHelper2::SetFlags() flags */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum DSH_FLAGS {
    DSH_ALLOWDROPDESCRIPTIONTEXT    = 0x0001
} DSH_FLAGS;
#endif

/* IShellLink::Resolve() flags */
typedef enum SLR_FLAGS {
    SLR_NO_UI                       = 0x0001,
    SLR_ANY_MATCH                   = 0x0002,
    SLR_UPDATE                      = 0x0004,
    SLR_NOUPDATE                    = 0x0008,
    SLR_NOSEARCH                    = 0x0010,
    SLR_NOTRACK                     = 0x0020,
    SLR_NOLINKINFO                  = 0x0040,
    SLR_INVOKE_MSI                  = 0x0080,
    SLR_NO_UI_WITH_MSG_PUMP         = 0x0101,
    SLR_OFFER_DELETE_WITHOUT_FILE   = 0x0200,
    SLR_KNOWNFOLDER                 = 0x0400,
    SLR_MACHINE_IN_LOCAL_TARGET     = 0x0800,
    SLR_UPDATE_MACHINE_AND_SID      = 0x1000
} SLR_FLAGS;

/* IShellLink::GetPath() flags */
typedef enum SLGP_FLAGS {
    SLGP_SHORTPATH          = 0x0001,
    SLGP_UNCPRIORITY        = 0x0002,
    SLGP_RAWPATH            = 0x0004,
    SLGP_RELATIVEPRIORITY   = 0x0008
} SLGP_FLAGS;

/* IActionProgressDialog::Initialize() flags */
enum _SPINITF {
    SPINITF_NORMAL      = 0x00000000,
    SPINITF_MODAL       = 0x00000001,
    SPINITF_NOMINIMIZE  = 0x00000008
};
typedef DWORD   SPINITF;

/* IActionProgress::Begin() flags */
enum _SPBEGINF {
    SPBEGINF_NORMAL             = 0x00000000,
    SPBEGINF_AUTOTIME           = 0x00000001,
    SPBEGINF_NOPROGRESSBAR      = 0x00000010,
    SPBEGINF_MARQUEEPROGRESS    = 0x00000020,
    SPBEGINF_NOCANCELBUTTON     = 0x00000040
};
typedef DWORD   SPBEGINF;

/* IActionProgress actions */
typedef enum _SPACTION {
    SPACTION_NONE               = 0,
    SPACTION_MOVING             = 1,
    SPACTION_COPYING            = 2,
    SPACTION_RECYCLING          = 3,
    SPACTION_APPLYINGATTRIBS    = 4,
    SPACTION_DOWNLOADING        = 5,
    SPACTION_SEARCHING_INTERNET = 6,
    SPACTION_CALCULATING        = 7,
    SPACTION_UPLOADING          = 8,
    SPACTION_SEARCHING_FILES    = 9,
    SPACTION_DELETING           = 10,
    SPACTION_RENAMING           = 11,
    SPACTION_FORMATTING         = 12,
    SPACTION_COPY_MOVING        = 13
} SPACTION;

/* IActionProgress text types */
typedef enum _SPTEXT {
    SPTEXT_ACTIONDESCRIPTION    = 1,
    SPTEXT_ACTIONDETAIL         = 2
} SPTEXT;

/* IShellPropSheetExt page identifiers */
enum _EXPPS {
    EXPPS_FILETYPES = 0x00000001
};
typedef UINT    EXPPS;

/* Desk band information */
#include <pshpack8.h>
typedef struct DESKBANDINFO {
    DWORD       dwMask;
    POINTL      ptMinSize;
    POINTL      ptMaxSize;
    POINTL      ptIntegral;
    POINTL      ptActual;
    WCHAR       wszTitle[256];
    DWORD       dwModeFlags;
    COLORREF    crBkgnd;
} DESKBANDINFO;
#include <poppack.h>

/* Desk band command identifiers */
enum tagDESKBANDCID {
    DBID_BANDINFOCHANGED    = 0,
    DBID_SHOWONLY           = 1,
    DBID_MAXIMIZEBAND       = 2,
    DBID_PUSHCHEVRON        = 3,
    DBID_DELAYINIT          = 4,
    DBID_FINISHINIT         = 5,
    DBID_SETWINDOWTHEME     = 6,
    DBID_PERMITAUTOHIDE     = 7
};

/* Thumb button flags */
typedef enum THUMBBUTTONFLAGS {
    THBF_ENABLED        = 0x00000000,
    THBF_DISABLED       = 0x00000001,
    THBF_DISMISSONCLICK = 0x00000002,
    THBF_NOBACKGROUND   = 0x00000004,
    THBF_HIDDEN         = 0x00000008,
    THBF_NONINTERACTIVE = 0x00000010
} THUMBBUTTONFLAGS;

/* Thumb button mask flags */
typedef enum THUMBBUTTONMASK {
    THB_BITMAP  = 0x00000001,
    THB_ICON    = 0x00000002,
    THB_TOOLTIP = 0x00000004,
    THB_FLAGS   = 0x00000008
} THUMBBUTTONMASK;

/* Thumb button */
#include <pshpack8.h>
typedef struct THUMBBUTTON {
    THUMBBUTTONMASK     dwMask;
    UINT                iId;
    UINT                iBitmap;
    HICON               hIcon;
    WCHAR               szTip[260];
    THUMBBUTTONFLAGS    dwFlags;
} THUMBBUTTON;
typedef THUMBBUTTON *LPTHUMBBUTTON;
#include <poppack.h>

/* Taskbar progress flags */
typedef enum TBPFLAG {
    TBPF_NOPROGRESS     = 0x00000000,
    TBPF_INDETERMINATE  = 0x00000001,
    TBPF_NORMAL         = 0x00000002,
    TBPF_ERROR          = 0x00000004,
    TBPF_PAUSED         = 0x00000008
} TBPFLAG;

/* ITaskbarList4::SetTabProperties() flags */
typedef enum STPFLAG {
    STPF_NONE                       = 0x00000000,
    STPF_USEAPPTHUMBNAILALWAYS      = 0x00000001,
    STPF_USEAPPTHUMBNAILWHENACTIVE  = 0x00000002,
    STPF_USEAPPPEEKALWAYS           = 0x00000004,
    STPF_USEAPPPEEKWHENACTIVE       = 0x00000008
} STPFLAG;

/* Explorer browser options */
#if (_WIN32_IE >= 0x0700)
typedef enum EXPLORER_BROWSER_OPTIONS {
    EBO_NONE                = 0x00000000,
    EBO_NAVIGATEONCE        = 0x00000001,
    EBO_SHOWFRAMES          = 0x00000002,
    EBO_ALWAYSNAVIGATE      = 0x00000004,
    EBO_NOTRAVELLOG         = 0x00000008,
    EBO_NOWRAPPERWINDOW     = 0x00000010,
    EBO_HTMLSHAREPOINTVIEW  = 0x00000020
} EXPLORER_BROWSER_OPTIONS;
#endif

/* Explorer browser fill flags */
#if (_WIN32_IE >= 0x0700)
typedef enum EXPLORER_BROWSER_FILL_FLAGS {
    EBF_NONE                    = 0x00000000,
    EBF_SELECTFROMDATAOBJECT    = 0x00000100,
    EBF_NODROPTARGET            = 0x00000200
} EXPLORER_BROWSER_FILL_FLAGS;
#endif

/* Operations progress dialog flags */
#if (_WIN32_IE >= 0x0700)
enum _OPPROGDLGF {
    OPPROGDLG_ENABLEPAUSE           = 0x00000080,
    OPPROGDLG_ALLOWUNDO             = 0x00000100,
    OPPROGDLG_DONTDISPLAYSOURCEPATH = 0x00000200,
    OPPROGDLG_DONTDISPLAYDESTPATH   = 0x00000400
};
typedef DWORD   OPPROGDLGF;
#endif

/* Progress dialog modes */
#if (_WIN32_IE >= 0x0700)
enum _PDMODE {
    PDM_DEFAULT         = 0x00000000,
    PDM_RUN             = 0x00000001,
    PDM_PREFLIGHT       = 0x00000002,
    PDM_UNDOING         = 0x00000004,
    PDM_ERRORSBLOCKING  = 0x00000008,
    PDM_INDETERMINATE   = 0x00000010
};
typedef DWORD   PDMODE;
#endif

/* Progress dialog operation status values */
#if (_WIN32_IE >= 0x0700)
typedef enum PDOPSTATUS {
    PDOPS_RUNNING   = 1,
    PDOPS_PAUSED    = 2,
    PDOPS_CANCELLED = 3,
    PDOPS_STOPPED   = 4,
    PDOPS_ERRORS    = 5
} PDOPSTATUS;
#endif

/* Namespace walk flags */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
typedef enum {
    NSWF_NONE_IMPLIES_ALL               = 0x00000001,
    NSWF_ONE_IMPLIES_ALL                = 0x00000002,
    NSWF_DONT_TRAVERSE_LINKS            = 0x00000004,
    NSWF_DONT_ACCUMULATE_RESULT         = 0x00000008,
    NSWF_TRAVERSE_STREAM_JUNCTIONS      = 0x00000010,
    NSWF_FILESYSTEM_ONLY                = 0x00000020,
    NSWF_SHOW_PROGRESS                  = 0x00000040,
    NSWF_FLAG_VIEWORDER                 = 0x00000080,
    NSWF_IGNORE_AUTOPLAY_HIDA           = 0x00000100,
    NSWF_ASYNC                          = 0x00000200,
    NSWF_DONT_RESOLVE_LINKS             = 0x00000400,
    NSWF_ACCUMULATE_FOLDERS             = 0x00000800,
    NSWF_DONT_SORT                      = 0x00001000,
    NSWF_USE_TRANSFER_MEDIUM            = 0x00002000,
    NSWF_DONT_TRAVERSE_STREAM_JUNCTIONS = 0x00004000
} NAMESPACEWALKFLAG;
#endif

/* Band site information */
#if (_WIN32_IE >= 0x0400)
#include <pshpack8.h>
typedef struct tagBANDSITEINFO {
    DWORD   dwMask;
    DWORD   dwState;
    DWORD   dwStyle;
} BANDSITEINFO;
#include <poppack.h>
#endif

/* Band site identifiers */
#if (_WIN32_IE >= 0x0400)
enum tagBANDSITECID {
    BSID_BANDADDED      = 0,
    BSID_BANDREMOVED    = 1
};
#endif

/* CD burning extension return values */
#if (NTDDI_VERSION >= 0x05010000)
enum tagCDBURNINGEXTENSIONRET {
    CDBE_RET_DEFAULT            = 0x00000000,
    CDBE_RET_DONTRUNOTHEREXTS   = 0x00000001,
    CDBE_RET_STOPWIZARD         = 0x00000002
};
#endif

/* CD burning extension actions */
#if (NTDDI_VERSION >= 0x05010000)
enum tagCDBURNINGEXTENSIONACTION {
    CDBE_TYPE_MUSIC = 0x00000001,
    CDBE_TYPE_DATA  = 0x00000002,
    CDBE_TYPE_ALL   = 0xFFFFFFFF
};
typedef DWORD   CDBE_ACTIONS;
#endif

/* Menu band handler identifiers */
#if (NTDDI_VERSION >= 0x05010000)
enum tagMENUBANDHANDLERCID {
    MBHANDCID_PIDLSELECT    = 0
};
#endif

/* IMenuPopup::OnSelect() type values */
#if (_WIN32_IE >= 0x0600)
enum tagMENUPOPUPSELECT {
    MPOS_EXECUTE        = 0,
    MPOS_FULLCANCEL     = 1,
    MPOS_CANCELLEVEL    = 2,
    MPOS_SELECTLEFT     = 3,
    MPOS_SELECTRIGHT    = 4,
    MPOS_CHILDTRACKING  = 5
};
#endif

/* IMenuPopup::Popup() flags */
#if (_WIN32_IE >= 0x0600)
enum tagMENUPOPUPFLAGS {
    MPPF_SETFOCUS       = 0x00000001,
    MPPF_INITIALSELECT  = 0x00000002,
    MPPF_NOANIMATE      = 0x00000004,
    MPPF_KEYBOARD       = 0x00000010,
    MPPF_REPOSITION     = 0x00000020,
    MPPF_FORCEZORDER    = 0x00000040,
    MPPF_FINALSELECT    = 0x00000080,
    MPPF_ALIGN_LEFT     = 0x02000000,
    MPPF_ALIGN_RIGHT    = 0x04000000,
    MPPF_TOP            = 0x20000000,
    MPPF_LEFT           = 0x40000000,
    MPPF_RIGHT          = 0x60000000,
    MPPF_BOTTOM         = 0x80000000,
    MPPF_POS_MASK       = 0xE0000000
};
typedef int MF_POPUPFLAGS;
#endif

/* File usage types */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum FILE_USAGE_TYPE {
    FUT_PLAYING = 0,
    FUT_EDITING = 1,
    FUT_GENERIC = 2
} FILE_USAGE_TYPE;
#endif

/* IFileDialogEvents overwrite response values */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum FDE_OVERWRITE_RESPONSE {
    FDEOR_DEFAULT   = 0x00000000,
    FDEOR_ACCEPT    = 0x00000001,
    FDEOR_REFUSE    = 0x00000002
} FDE_OVERWRITE_RESPONSE;
#endif

/* IFileDialogEvents share violation response values */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum FDE_SHAREVIOLATION_RESPONSE {
    FDESVR_DEFAULT  = 0x00000000,
    FDESVR_ACCEPT   = 0x00000001,
    FDESVR_REFUSE   = 0x00000002
} FDE_SHAREVIOLATION_RESPONSE;
#endif

/* IFileDialog::AddPlace() flags */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum FDAP {
    FDAP_BOTTOM = 0x00000000,
    FDAP_TOP    = 0x00000001
} FDAP;
#endif

/* File open dialog options */
#if (NTDDI_VERSION >= 0x06000000)
enum _FILEOPENDIALOGOPTIONS {
    FOS_OVERWRITEPROMPT     = 0x00000002,
    FOS_STRICTFILETYPES     = 0x00000004,
    FOS_NOCHANGEDIR         = 0x00000008,
    FOS_PICKFOLDERS         = 0x00000020,
    FOS_FORCEFILESYSTEM     = 0x00000040,
    FOS_ALLNONSTORAGEITEMS  = 0x00000080,
    FOS_NOVALIDATE          = 0x00000100,
    FOS_ALLOWMULTISELECT    = 0x00000200,
    FOS_PATHMUSTEXIST       = 0x00000800,
    FOS_FILEMUSTEXIST       = 0x00001000,
    FOS_CREATEPROMPT        = 0x00002000,
    FOS_SHAREAWARE          = 0x00004000,
    FOS_NOREADONLYRETURN    = 0x00008000,
    FOS_NOTESTFILECREATE    = 0x00010000,
    FOS_HIDEMRUPLACES       = 0x00020000,
    FOS_HIDEPINNEDPLACES    = 0x00040000,
    FOS_NODEREFERENCELINKS  = 0x00100000,
    FOS_DONTADDTORECENT     = 0x02000000,
    FOS_FORCESHOWHIDDEN     = 0x10000000,
    FOS_DEFAULTNOMINIMODE   = 0x20000000,
    FOS_FORCEPREVIEWPANEON  = 0x40000000
};
typedef DWORD   FILEOPENDIALOGOPTIONS;
#endif

/* CD control state */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum CDCONTROLSTATEF {
    CDCS_INACTIVE       = 0x00000000,
    CDCS_ENABLED        = 0x00000001,
    CDCS_VISIBLE        = 0x00000002,
    CDCS_ENABLEDVISIBLE = 0x00000003
} CDCONTROLSTATEF;
#endif

/* Association levels */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum ASSOCIATIONLEVEL {
    AL_MACHINE      = 0,
    AL_EFFECTIVE    = 1,
    AL_USER         = 2
} ASSOCIATIONLEVEL;
#endif

/* Association types */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum ASSOCIATIONTYPE {
    AT_FILEEXTENSION    = 0,
    AL_URLPROTOCOL      = 1,
    AL_STARTMENUCLIENT  = 2,
    AT_MIMETYPE         = 3
} ASSOCIATIONTYPE;
#endif

/* Delegate item identifier */
#include <pshpack1.h>
typedef struct DELEGATEITEMID {
    WORD    cbSize;
    WORD    wOuter;
    WORD    cbInner;
    BYTE    rgb[1];
} DELEGATEITEMID;
typedef UNALIGNED DELEGATEITEMID        *PDELEGATEITEMID;
typedef const UNALIGNED DELEGATEITEMID  *PCDELEGATEITEMID;
#include <poppack.h>

/* Browser frame options */
#if (_WIN32_IE >= 0x0600)
enum _BROWSERFRAMEOPTIONS {
    BFO_NONE                                = 0x00000000,
    BFO_BROWSER_PERSIST_SETTINGS            = 0x00000001,
    BFO_RENAME_FOLDER_OPTIONS_TOINTERNET    = 0x00000002,
    BFO_BOTH_OPTIONS                        = 0x00000004,
    BIF_PREFER_INTERNET_SHORTCUT            = 0x00000008,
    BFO_BROWSE_NO_IN_NEW_PROCESS            = 0x00000010,
    BFO_ENABLE_HYPERLINK_TRACKING           = 0x00000020,
    BFO_USE_IE_OFFLINE_SUPPORT              = 0x00000040,
    BFO_SUBSTITUTE_INTERNET_START_PAGE      = 0x00000080,
    BFO_USE_IE_LOGOBANDING                  = 0x00000100,
    BFO_ADD_IE_TOCAPTIONBAR                 = 0x00000200,
    BFO_USE_DIALUP_REF                      = 0x00000400,
    BFO_USE_IE_TOOLBAR                      = 0x00000800,
    BFO_NO_PARENT_FOLDER_SUPPORT            = 0x00001000,
    BFO_NO_REOPEN_NEXT_RESTART              = 0x00002000,
    BFO_GO_HOME_PAGE                        = 0x00004000,
    BFO_PREFER_IEPROCESS                    = 0x00008000,
    BFO_SHOW_NAVIGATION_CANCELLED           = 0x00010000,
    BFO_USE_IE_STATUSBAR                    = 0x00020000,
    BFO_QUERY_ALL                           = 0xFFFFFFFF
};
typedef DWORD   BROWSERFRAMEOPTIONS;
#endif

/* INewWindowManager flags */
#if (_WIN32_IE >= 0x0602)
typedef enum NWMF {
    NWMF_UNLOADING          = 0x00000001,
    NWMF_USERINITED         = 0x00000002,
    NWMF_FIRST              = 0x00000004,
    NWMF_OVERRIDEKEY        = 0x00000008,
    NWMF_SHOWHELP           = 0x00000010,
    NWMF_HTMLDIALOG         = 0x00000020,
    NWMF_FROMDIALOGCHILD    = 0x00000040,
    NWMF_USERREQUESTED      = 0x00000080,
    NWMF_USERALLOWED        = 0x00000100,
    NWMF_FORCEWINDOW        = 0x00010000,
    NWMF_FORCETAB           = 0x00020000,
    NWMF_SUGGESTWINDOW      = 0x00040000,
    NWMF_SUGGESTTAB         = 0x00080000,
    NWMF_INACTIVETAB        = 0x00100000
} NWMF;
#endif

/* Attachment prompts */
#if (_WIN32_IE >= 0x0602)
typedef enum ATTACHMENT_PROMPT {
    ATTACHMENT_PROMPT_NONE          = 0x00000000,
    ATTACHMENT_PROMPT_SAVE          = 0x00000001,
    ATTACHMENT_PROMPT_EXEC          = 0x00000002,
    ATTACHMENT_PROMPT_EXEC_OR_SAVE  = 0x00000003
} ATTACHMENT_PROMPT;
#endif

/* Attachment actions */
#if (_WIN32_IE >= 0x0602)
typedef enum ATTACHMENT_ACTION {
    ATTACHMENT_ACTION_CANCEL    = 0x00000000,
    ATTACHMENT_ACTION_SAVE      = 0x00000001,
    ATTACHMENT_ACTION_EXEC      = 0x00000002
} ATTACHMENT_ACTION;
#endif

/* Shell menu data */
#if (_WIN32_IE >= 0x0600)
#include <pshpack8.h>
typedef struct tagSMDATA {
    DWORD               dwMask;
    DWORD               dwFlags;
    HMENU               hmenu;
    HWND                hwnd;
    UINT                uId;
    UINT                uIdParent;
    UINT                uIdAncestor;
    IUnknown            *punk;
    PIDLIST_ABSOLUTE    pidlFolder;
    PUITEMID_CHILD      pidlItem;
    IShellFolder        *psf;
    void                *pvUserData;
} SMDATA;
typedef SMDATA  *LPSMDATA;
#include <poppack.h>
#endif

/* Shell menu information */
#if (_WIN32_IE >= 0x0600)
#include <pshpack8.h>
typedef struct tagSMINFO {
    DWORD   dwMask;
    DWORD   dwType;
    DWORD   dwFlags;
    int     iIcon;
} SMINFO;
typedef SMINFO  *PSMINFO;
#include <poppack.h>
#endif

/* IShellMenuCallback change notification structure */
#if (_WIN32_IE >= 0x0600)
#include <pshpack8.h>
typedef struct SMCSHCHANGENOTIFYSTRUCT {
    long                lEvent;
    PCIDLIST_ABSOLUTE   pidl1;
    PCIDLIST_ABSOLUTE   pidl2;
} SMCSHCHANGENOTIFYSTRUCT;
typedef SMCSHCHANGENOTIFYSTRUCT *PSMCSHCHANGENOTIFYSTRUCT;
#include <poppack.h>
#endif

/* Shell menu information masks */
#if (_WIN32_IE >= 0x0600)
enum tagSMINFOMASK {
    SMIM_TYPE   = 0x00000001,
    SMIM_FLAGS  = 0x00000002,
    SMIM_ICON   = 0x00000004
};
#endif

/* Shell menu information types */
#if (_WIN32_IE >= 0x0600)
enum tagSMINFOTYPE {
    SMIT_SEPARATOR  = 0x00000001,
    SMIT_STRING     = 0x00000002
};
#endif

/* Shell menu information flags */
#if (_WIN32_IE >= 0x0600)
enum tagSMINFOFLAGS {
    SMIF_ICON           = 0x00000001,
    SMIF_ACCELERATOR    = 0x00000002,
    SMIF_DROPTARGET     = 0x00000004,
    SMIF_SUBMENU        = 0x00000008,
    SMIF_CHECKED        = 0x00000020,
    SMIF_DROPCASCADE    = 0x00000040,
    SMIF_HIDDEN         = 0x00000080,
    SMIF_DISABLED       = 0x00000100,
    SMIF_TRACKPOPUP     = 0x00000200,
    SMIF_DEMOTED        = 0x00000400,
    SMIF_ALTSTATE       = 0x00000800,
    SMIF_DRAGNDROP      = 0x00001000,
    SMIF_NEW            = 0x00002000
};
#endif

/* Known folder categories */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum KF_CATEGORY {
    KF_CATEGORY_VIRTUAL = 0x00000001,
    KF_CATEGORY_FIXED   = 0x00000002,
    KF_CATEGORY_COMMON  = 0x00000003,
    KF_CATEGORY_PERUSER = 0x00000004
} KF_CATEGORY;
#endif

/* Known folder definition flags */
#if (NTDDI_VERSION >= 0x06000000)
enum _KF_DEFINITION_FLAGS {
    KFDF_LOCAL_REDIRECT_ONLY    = 0x00000002,
    KFDF_ROAMABLE               = 0x00000004,
    KFDF_PRECREATE              = 0x00000008,
    KFDF_STREAM                 = 0x00000010,
    KFDF_PUBLISHEXPANDEDPATH    = 0x00000020
};
typedef DWORD   KF_DEFINITION_FLAGS;
#endif

/* Known folder redirect flags */
#if (NTDDI_VERSION >= 0x06000000)
enum _KF_REDIRECT_FLAGS {
    KF_REDIRECT_USER_EXCLUSIVE                  = 0x00000001,
    KF_REDIRECT_COPY_SOURCE_DACL                = 0x00000002,
    KF_REDIRECT_OWNER_USER                      = 0x00000004,
    KF_REDIRECT_SET_OWNER_EXPLICIT              = 0x00000008,
    KF_REDIRECT_CHECK_ONLY                      = 0x00000010,
    KF_REDIRECT_WITH_UI                         = 0x00000020,
    KF_REDIRECT_UNPIN                           = 0x00000040,
    KF_REDIRECT_PIN                             = 0x00000080,
    KF_REDIRECT_COPY_CONTENTS                   = 0x00000200,
    KF_REDIRECT_DEL_SOURCE_CONTENTS             = 0x00000400,
    KF_REDIRECT_EXCLUDE_ALL_KNOWN_SUBFOLDERS    = 0x00000800
};
typedef DWORD KF_REDIRECT_FLAGS;
#endif

/* Known folder redirection capabilities */
#if (NTDDI_VERSION >= 0x06000000)
enum _KF_REDIRECTION_CAPABILITIES {
    KF_REDIRECTION_CAPABILITIES_ALLOW_ALL               = 0x000000FF,
    KF_REDIRECTION_CAPABILITIES_REDIRECTABLE            = 0x00000001,
    KF_REDIRECTION_CAPABILITIES_DENY_ALL                = 0x000FFF00,
    KF_REDIRECTION_CAPABILITIES_DENY_POLICY_REDIRECTED  = 0x00000100,
    KF_REDIRECTION_CAPABILITIES_DENY_POLICY             = 0x00000200,
    KF_REDIRECTION_CAPABILITIES_DENY_PERMISSIONS        = 0x00000400
};
typedef DWORD   KF_REDIRECTION_CAPABILITIES;
#endif

/* Known folder definition */
#if (NTDDI_VERSION >= 0x06000000)
typedef struct KNOWNFOLDER_DEFINITION {
    KF_CATEGORY         category;
    LPWSTR              pszName;
    LPWSTR              pszDescription;
    KNOWNFOLDERID       fidParent;
    LPWSTR              pszRelativePath;
    LPWSTR              pszParsingName;
    LPWSTR              pszTooltip;
    LPWSTR              pszLocalizedName;
    LPWSTR              pszIcon;
    LPWSTR              pszSecurity;
    DWORD               dwAttributes;
    KF_DEFINITION_FLAGS kfdFlags;
    FOLDERTYPEID        ftidType;
} KNOWNFOLDER_DEFINITION;
#endif

/* IKnownFolderManager::FindFolderFromPath() modes */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum FFFP_MODE {
    FFFP_EXACTMATCH         = 0,
    FFFP_NEARESTPARENTMATCH = 1
} FFFP_MODE;
#endif

/* Share roles */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum SHARE_ROLE {
    SHARE_ROLE_INVALID      = -1,
    SHARE_ROLE_READER       = 0,
    SHARE_ROLE_CONTRIBUTOR  = 1,
    SHARE_ROLE_CO_OWNER     = 2,
    SHARE_ROLE_OWNER        = 3,
    SHARE_ROLE_CUSTOM       = 4,
    SHARE_ROLE_MIXED        = 5
} SHARE_ROLE;
#endif

/* Share identifiers */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum DEF_SHARE_ID {
    DEFSHAREID_USERS    = 1,
    DEFSHAREID_PUBLIC   = 2
} DEF_SHARE_ID;
#endif

/* INewMenuClient::IncludeItems() flags */
enum _NMCII_FLAGS {
    NMCII_ITEMS     = 0x00000001,
    NMCII_FOLDERS   = 0x00000002
};
typedef int NMCII_FLAGS;

/* INewMenuClient::SelectAndEditItem() flags */
enum _NMCSAEI_FLAGS {
    NMCSAEI_SELECT  = 0x00000000,
    NMCSAEI_EDIT    = 0x00000001
};
typedef int NMCSAEI_FLAGS;

/* INameSpaceTreeControl styles */
enum _NSTCSTYLE {
    NSTCS_HASEXPANDOS           = 0x00000001,
    NSTCS_HASLINES              = 0x00000002,
    NSTCS_SINGLECLICKEXPAND     = 0x00000004,
    NSTCS_FULLROWSELECT         = 0x00000008,
    NSTCS_SPRINGEXPAND          = 0x00000010,
    NSTCS_HORIZONTALSCROLL      = 0x00000020,
    NSTCS_ROOTHASEXPANDO        = 0x00000040,
    NSTCS_SHOWSELECTIONALWAYS   = 0x00000080,
    NSTCS_NOINFOTIP             = 0x00000200,
    NSTCS_EVENHEIGHT            = 0x00000400,
    NSTCS_NOREPLACEOPEN         = 0x00000800,
    NSTCS_DISABLEDRAGDROP       = 0x00001000,
    NSTCS_NOORDERSTREAM         = 0x00002000,
    NSTCS_RICHTOOLTIP           = 0x00004000,
    NSTCS_BORDER                = 0x00008000,
    NSTCS_NOEDITLABELS          = 0x00010000,
    NSTCS_TABSTOP               = 0x00020000,
    NSTCS_FAVORITESMODE         = 0x00080000,
    NSTCS_AUTOHSCROLL           = 0x00100000,
    NSTCS_FADEINOUTEXPANDOS     = 0x00200000,
    NSTCS_EMPTYTEXT             = 0x00400000,
    NSTCS_CHECKBOXES            = 0x00800000,
    NSTCS_PARTIALCHECKBOXES     = 0x01000000,
    NSTCS_EXCLUSIONCHECKBOXES   = 0x02000000,
    NSTCS_DIMMENDCHECKBOXES     = 0x04000000,
    NSTCS_NOINDENTCHECKS        = 0x08000000,
    NSTCS_ALLOWJUNCTIONS        = 0x10000000,
    NSTCS_SHOWTABSBUTTON        = 0x20000000,
    NSTCS_SHOWDELETEBUTTON      = 0x40000000,
    NSTCS_SHOWREFRESHBUTTON     = 0x80000000
};
typedef DWORD   NSTCSTYLE;

/* INameSpaceTreeControl root styles */
enum _NSTCROOTSTYLE {
    NSTCRS_VISIBLE  = 0x00000000,
    NSTCRS_HIDDEN   = 0x00000001,
    NSTCRS_EXPANDED = 0x00000002
};
typedef DWORD   NSTCROOTSTYLE;

/* INameSpaceTreeControl item states */
enum _NSTCITEMSTATE {
    NSTCIS_NONE             = 0x00000000,
    NSTCIS_SELECTED         = 0x00000001,
    NSTCIS_EXPANDED         = 0x00000002,
    NSTCIS_BOLD             = 0x00000004,
    NSTCIS_DISABLED         = 0x00000008,
    NCTCIS_SELECTEDNOEXPAND = 0x00000010
};
typedef DWORD   NSTCITEMSTATE;

/* INameSpaceTreeControl::GetNextItem() codes */
typedef enum NSTCGNI {
    NSTCGNI_NEXT            = 0,
    NSTCGNI_NEXTVISIBLE     = 1,
    NSTCGNI_PREV            = 2,
    NSTCGNI_PREVVISIBLE     = 3,
    NSTCGNI_PARENT          = 4,
    NSTCGNI_CHILD           = 5,
    NSTCGNI_FIRSTVISIBLE    = 6,
    NSTCGNI_LASTVISIBLE     = 7
} NSTCGNI;

/* INameSpaceTreeControl2 styles */
typedef enum NSTCSTYLE2 {
    NSTCS2_DEFAULT                  = 0x00000000,
    NSTCS2_INTERRUPTNOTIFICATIONS   = 0x00000001,
    NSTCS2_SHOWNULLSPACEMENU        = 0x00000002,
    NSTCS2_DISPLAYPADDING           = 0x00000004,
    NSTCS2_DISPLAYPINNEDONLY        = 0x00000008,
    NSTCS2_NOSINGLETONAUTOEXPAND    = 0x00000010,
    NSTCS2_NEVERINSERTNONENUMERATED = 0x00000020
} NSTCSTYLE2;

/* INameSpaceTreeControlEvents hit test values */
enum _NSTCEHITTEST {
    NSTCEHT_NOWHERE         = 0x00000001,
    NSTCEHT_ONITEMICON      = 0x00000002,
    NSTCEHT_ONITEMLABEL     = 0x00000004,
    NSTCEHT_ONITEMINDENT    = 0x00000008,
    NSTCEHT_ONITEMBUTTON    = 0x00000010,
    NSTCEHT_ONITEMRIGHT     = 0x00000020,
    NSTCEHT_ONITEMSTATEICON = 0x00000040,
    NSTCEHT_ONITEM          = 0x00000046,
    NSTCEHT_ONITEMTABBUTTON = 0x00001000
};
typedef DWORD   NSTCEHITTEST;

/* INameSpaceTreeControlEvents click types */
enum _NSTCECLICKTYPE {
    NSTCECT_LBUTTON     = 0x00000001,
    NSTCECT_MBUTTON     = 0x00000002,
    NSTCECT_RBUTTON     = 0x00000003,
    NSTCECT_BUTTON      = 0x00000003,
    NSTCECT_DBLCLICK    = 0x00000004
};
typedef DWORD   NSTCECLICKTYPE;

/* Macros to manipulate click types */
#define ISLBUTTON( x )  (((x) & NSTCECT_BUTTON) == NSTCECT_LBUTTON)
#define ISMBUTTON( x )  (((x) & NSTCECT_BUTTON) == NSTCECT_MBUTTON)
#define ISRBUTTON( x )  (((x) & NSTCECT_BUTTON) == NSTCECT_RBUTTON)
#define ISDBLCLICK( x ) (((x) & NSTCECT_DBLCLICK) == NSTCECT_DBLCLICK)

/* INameSpaceTreeControl custom draw parameters */
typedef struct NSTCCUSTOMDRAW {
    IShellItem      *psi;
    UINT            uItemState;
    NSTCITEMSTATE   nstcis;
    LPCWSTR         pszText;
    int             iImage;
    HIMAGELIST      himl;
    int             iLevel;
    int             iIndent;
} NSTCCUSTOMDRAW;

/* INameSpaceTreeControlFolderCapabilities flags */
#if (NTDDI_VERSION >= 0x06000000)
typedef enum NSTCFOLDERCAPABILITIES {
    NSTCFC_NONE                     = 0x00000000,
    NSTCFC_PINNEDITEMFILTERING      = 0x00000001,
    NSTCFC_DELAY_REGISTER_NOTIFY    = 0x00000002
} NSTCFOLDERCAPABILITIES;
#endif

/* Preview handler frame information */
typedef struct {
    HACCEL  haccel;
    UINT    cAccelEntries;
} PREVIEWHANDLERFRAMEINFO;

/* Explorer pane states */
#if (NTDDI_VERSION >= 0x06000000)
enum _EXPLORERPANESTATE {
    EPS_DONTCARE        = 0x00000000,
    EPS_DEFAULT_ON      = 0x00000001,
    EPS_DEFAULT_OFF     = 0x00000002,
    EPS_STATEMASK       = 0x0000FFFF,
    EPS_INITIALSTATE    = 0x00010000,
    EPS_FORCE           = 0x00020000
};
typedef DWORD   EXPLORERPANESTATE;
#endif

/* Explorer command states */
enum _EXPCMDSTATE {
    ECS_ENABLED     = 0x00000000,
    ECS_DISABLED    = 0x00000001,
    ECS_HIDDEN      = 0x00000002,
    ECS_CHECKBOX    = 0x00000004,
    ECS_CHECKED     = 0x00000008,
    ECS_RADIOCHECK  = 0x00000010
};
typedef DWORD   EXPCMDSTATE;

/* Explorer command flags */
enum _EXPCMDFLAGS {
    ECF_HASSUBCOMMANDS  = 0x00000001,
    ECF_HASSPLITBUTTON  = 0x00000002,
    ECF_HIDELABEL       = 0x00000004,
    ECF_ISSEPARATOR     = 0x00000008,
    ECF_HASLUASHIELD    = 0x00000010,
    ECF_SEPARATORBEFORE = 0x00000020,
    ECF_SEPARATORAFTER  = 0x00000040,
    ECF_ISDROPDOWN      = 0x00000080
};
typedef DWORD   EXPCMDFLAGS;

/* Markup size values */
typedef enum tagMARKUPSIZE {
    MARKUPSIZE_CALCWIDTH    = 0,
    MARKUPSIZE_CALCHEIGHT   = 1
} MARKUPSIZE;

/* Markup link text values */
typedef enum tagMARKUPLINKTEXT {
    MARKUPLINKTEXT_URL  = 0,
    MARKUPLINKTEXT_ID   = 1,
    MARKUPLINKTEXT_TEXT = 2
} MARKUPLINKTEXT;

/* Markup states */
enum tagMARKUPSTATE {
    MARKUPSTATE_FOCUSED         = 0x00000001,
    MARKUPSTATE_ENABLED         = 0x00000002,
    MARKUPSTATE_VISITED         = 0x00000004,
    MARKUPSTATE_HOT             = 0x00000008,
    MARKUPSTATE_DEFAULTCOLORS   = 0x00000010,
    MARKUPSTATE_ALLOWMARKUP     = 0x40000000
};
typedef DWORD   MARKUPSTATE;

/* Markup messages */
typedef enum tagMARKUPMESSAGE {
    MARKUPMESSAGE_KEYEXECUTE    = 0,
    MARKUPMESSAGE_CLICKEXECUTE  = 1,
    MARKUPMESSAGE_WANTFOCUS     = 2
} MARKUPMESSAGE;

/* Control panel views */
typedef enum CPVIEW {
    CPVIEW_CLASSIC  = 0,
    CPVIEW_ALLITEMS = CPVIEW_CLASSIC,
    CPVIEW_CATEGORY = 1,
    CPVIEW_HOME     = CPVIEW_CATEGORY
} CPVIEW;

/* Known destination categories */
#if (NTDDI_VERSION >= 0x06010000)
typedef enum KNOWNDESTCATEGORY {
    KDC_FREQUENT    = 1,
    KDC_RECENT      = 2
} KNOWNDESTCATEGORY;
#endif

/* Application document list types */
#if (NTDDI_VERSION >= 0x06010000)
typedef enum APPDOCLISTTYPE {
    ADLT_RECENT     = 0,
    ADLT_FREQUENT   = 1
} APPDOCLISTTYPE;
#endif

/* HomeGroup sharing choices */
typedef enum HOMEGROUPSHARINGCHOICES {
    HGSC_NONE               = 0x00000000,
    HGSC_MUSICLIBRARY       = 0x00000001,
    HGSC_PICTURESLIBRARY    = 0x00000002,
    HGSC_VIDEOSLIBRARY      = 0x00000004,
    HGSC_DOCUMENTSLIBRARY   = 0x00000008,
    HGSC_PRINTERS           = 0x00000010
} HOMEGROUPSHARINGCHOICES;

/* Library folder filters */
typedef enum LIBRARYFOLDERFILTER {
    LFF_FORCEFILESYSTEM = 1,
    LFF_STORAGEITEMS    = 2,
    LFF_ALLITEMS        = 3
} LIBRARYFOLDERFILTER;

/* Library option flags */
typedef enum LIBRARYOPTIONFLAGS {
    LOF_DEFAULT         = 0x00000000,
    LOF_PINNEDTONAVPANE = 0x00000001,
    LOF_MASK_ALL        = 0x00000001
} LIBRARYOPTIONFLAGS;

/* Default save folder types */
typedef enum DEFAULTSAVEFOLDERTYPE {
    DSFT_DETECT     = 1,
    DSFT_PRIVATE    = 2,
    DSFT_PUBLIC     = 3
} DEFAULTSAVEFOLDERTYPE;

/* Library save flags */
typedef enum LIBRARYSAVEFLAGS {
    LSF_FAILIFTHERE         = 0x00000000,
    LSF_OVERRIDEEXISTING    = 0x00000001,
    LSF_MAKEUNIQUENAME      = 0x00000002
} LIBRARYSAVEFLAGS;

/* Library manage dialog options */
#if (NTDDI_VERSION >= 0x06010000) && (_WIN32_IE >= 0x0700)
typedef enum LIBRARYMANAGEDIALOGOPTIONS {
    LMD_DEFAULT                             = 0x00000000,
    LMD_ALLOWUNINDEXABLENETWORKLOCATIONS    = 0x00000001
} LIBRARYMANAGEDIALOGOPTIONS;
#endif

/* Association filters */
#if (NTDDI_VERSION >= 0x06000000)
enum _ASSOC_FILTER {
    ASSOC_FILTER_NONE           = 0x00000000,
    ASSOC_FILTER_RECOMMENDED    = 0x00000001
};
typedef int ASSOC_FILTER;
#endif
    
/* GUIDs */
EXTERN_C const IID      IID_IContextMenu;
EXTERN_C const IID      IID_IContextMenu2;
EXTERN_C const IID      IID_IContextMenu3;
EXTERN_C const IID      IID_IExecuteCommand;
EXTERN_C const IID      IID_IPersistFolder;
#if (_WIN32_IE >= 0x0400)
EXTERN_C const IID      IID_IRunnableTask;
EXTERN_C const IID      IID_IShellTaskScheduler;
EXTERN_C const IID      IID_IQueryCodePage;
EXTERN_C const IID      IID_IPersistFolder2;
#endif
#if (_WIN32_IE >= 0x0500)
EXTERN_C const IID      IID_IPersistFolder3;
#endif
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0600)
EXTERN_C const IID      IID_IPersistIDList;
#endif
EXTERN_C const IID      IID_IEnumIDList;
EXTERN_C const IID      IID_IEnumFullIDList;
#if (NTDDI_VERSION >= 0x06010000)
EXTERN_C const IID      IID_IObjectWithFolderEnumMode;
EXTERN_C const IID      IID_IParseAndCreateItem;
#endif
EXTERN_C const IID      IID_IShellFolder;
EXTERN_C const IID      IID_IEnumExtraSearch;
EXTERN_C const IID      IID_IShellFolder2;
EXTERN_C const IID      IID_IFolderViewOptions;
EXTERN_C const IID      IID_IShellView;
EXTERN_C const IID      IID_IShellView2;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IShellView3;
#endif
EXTERN_C const IID      IID_IFolderView;
#if (NTDDI_VERSION >= 0x06010000)
EXTERN_C const IID      IID_ISearchBoxInfo;
#endif
#if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IFolderView2;
#endif
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IFolderViewSettings;
#endif
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IPreviewHandlerVisuals;
EXTERN_C const IID      IID_IVisualProperties;
#endif
EXTERN_C const IID      IID_ICommDlgBrowser;
#if (NTDDI_VERSION >= 0x05000000)
EXTERN_C const IID      IID_ICommDlgBrowser2;
#endif
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_ICommDlgBrowser3;
EXTERN_C const IID      IID_IColumnManager;
#endif
EXTERN_C const IID      IID_IFolderFilterSite;
EXTERN_C const IID      IID_IFolderFilter;
EXTERN_C const IID      IID_IInputObjectSite;
EXTERN_C const IID      IID_IInputObject;
EXTERN_C const IID      IID_IInputObject2;
EXTERN_C const IID      IID_IShellIcon;
EXTERN_C const IID      IID_IShellBrowser;
EXTERN_C const IID      IID_IProfferService;
EXTERN_C const IID      IID_IShellItem;
EXTERN_C const IID      IID_IShellItem2;
EXTERN_C const IID      IID_IShellItemImageFactory;
EXTERN_C const IID      IID_IUserAccountChangeCallback;
#if (NTDDI_VERSION >= 0x05010000)
EXTERN_C const IID      IID_IEnumShellItems;
#endif
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_ITransferAdviseSink;
#endif
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_ITransferSource;
#endif
EXTERN_C const IID      IID_IEnumResources;
EXTERN_C const IID      IID_IShellItemResources;
EXTERN_C const IID      IID_ITransferDestination;
EXTERN_C const IID      IID_IStreamAsync;
EXTERN_C const IID      IID_IStreamUnbufferedInfo;
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IFileOperationProgressSink;
#endif
EXTERN_C const IID      IID_IShellItemArray;
EXTERN_C const IID      IID_IInitializeWithItem;
EXTERN_C const IID      IID_IObjectWithSelection;
EXTERN_C const IID      IID_IObjectWithBackReferences;
EXTERN_C const IID      IID_IPropertyUI;
#if (_WIN32_IE >= 0x0500)
EXTERN_C const IID      IID_ICategoryProvider;
EXTERN_C const IID      IID_ICategorizer;
#endif
#if (NTDDI_VERSION >= 0x05000000)
EXTERN_C const IID      IID_IDropTargetHelper;
EXTERN_C const IID      IID_IDragSourceHelper;
#endif
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IDragSourceHelper2;
#endif
EXTERN_C const IID      IID_IShellLinkA;
EXTERN_C const IID      IID_IShellLinkW;
EXTERN_C const IID      IID_IShellLinkDataList;
#if (NTDDI_VERSION >= 0x05000000)
EXTERN_C const IID      IID_IResolveShellLink;
#endif
EXTERN_C const IID      IID_IActionProgressDialog;
EXTERN_C const IID      IID_IHWEventHandler;
EXTERN_C const IID      IID_IHWEventHandler2;
EXTERN_C const IID      IID_IQueryCancelAutoPlay;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IDynamicHWHandler;
#endif
EXTERN_C const IID      IID_IActionProgress;
EXTERN_C const IID      IID_IShellExtInit;
EXTERN_C const IID      IID_IShellPropSheetExt;
EXTERN_C const IID      IID_IRemoteComputer;
EXTERN_C const IID      IID_IQueryContinue;
EXTERN_C const IID      IID_IObjectWithCancelEvent;
EXTERN_C const IID      IID_IUserNotification;
EXTERN_C const IID      IID_IUserNotificationCallback;
EXTERN_C const IID      IID_IUserNotification2;
EXTERN_C const IID      IID_IItemNameLimits;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_ISearchFolderItemFactory;
#endif
#if (_WIN32_IE >= 0x0400)
EXTERN_C const IID      IID_IExtractImage;
#endif
#if (_WIN32_IE >= 0x0500)
EXTERN_C const IID      IID_IExtractImage2;
EXTERN_C const IID      IID_IThumbnailHandlerFactory;
EXTERN_C const IID      IID_IParentAndItem;
#endif
EXTERN_C const IID      IID_IDockingWindow;
EXTERN_C const IID      IID_IDeskBand;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IDeskBandInfo;
EXTERN_C const IID      IID_IDeskBand2;
#endif
EXTERN_C const IID      IID_ITaskbarList;
EXTERN_C const IID      IID_ITaskbarList2;
EXTERN_C const IID      IID_ITaskbarList3;
EXTERN_C const IID      IID_ITaskbarList4;
EXTERN_C const IID      IID_IStartMenuPinnedList;
EXTERN_C const IID      IID_ICDBurn;
EXTERN_C const IID      IID_IWizardSite;
EXTERN_C const IID      IID_IWizardExtension;
EXTERN_C const IID      IID_IWebWizardExtension;
EXTERN_C const IID      IID_IPublishingWizard;
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IFolderViewHost;
#endif
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IExplorerBrowserEvents;
EXTERN_C const IID      IID_IExplorerBrowser;
EXTERN_C const IID      IID_IAccessibleObject;
#endif
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IResultsFolder;
#endif
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IEnumObjects;
EXTERN_C const IID      IID_IOperationsProgressDialog;
EXTERN_C const IID      IID_IIOCancelInformation;
EXTERN_C const IID      IID_IFileOperation;
EXTERN_C const IID      IID_IObjectProvider;
#endif
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_INamespaceWalkCB;
#endif
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_INamespaceWalkCB2;
#endif
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_INamespaceWalk;
#endif
EXTERN_C const IID      IID_IAutoCompleteDropDown;
#if (_WIN32_IE >= 0x0400)
EXTERN_C const IID      IID_IBandSite;
#endif
#if (NTDDI_VERSION >= 0x05010000)
EXTERN_C const IID      IID_IModalWindow;
#endif
EXTERN_C const IID      IID_IContextMenuSite;
EXTERN_C const IID      IID_IEnumReadyCallback;
EXTERN_C const IID      IID_IEnumerableView;
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IInsertItem;
#endif
#if (NTDDI_VERSION >= 0x05010000)
EXTERN_C const IID      IID_IMenuBand;
EXTERN_C const IID      IID_IFolderBandPriv;
EXTERN_C const IID      IID_IRegTreeItem;
EXTERN_C const IID      IID_IImageRecompress;
#endif
#if (_WIN32_IE >= 0x0600)
EXTERN_C const IID      IID_IDeskBar;
EXTERN_C const IID      IID_IMenuPopup;
#endif
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IFileIsInUse;
EXTERN_C const IID      IID_IFileDialogEvents;
EXTERN_C const IID      IID_IFileDialog;
EXTERN_C const IID      IID_IFileSaveDialog;
EXTERN_C const IID      IID_IFileOpenDialog;
EXTERN_C const IID      IID_IFileDialogCustomize;
EXTERN_C const IID      IID_IFileDialogControlEvents;
EXTERN_C const IID      IID_IFileDialog2;
EXTERN_C const IID      IID_IApplicationAssociationRegistration;
EXTERN_C const IID      IID_IApplicationAssociationRegistrationUI;
#endif
EXTERN_C const IID      IID_IDelegateFolder;
#if (_WIN32_IE >= 0x0600)
EXTERN_C const IID      IID_IBrowserFrameOptions;
#endif
#if (_WIN32_IE >= 0x0602)
EXTERN_C const IID      IID_INewWindowManager;
EXTERN_C const IID      IID_IAttachmentExecute;
#endif
#if (_WIN32_IE >= 0x0600)
EXTERN_C const IID      IID_IShellMenuCallback;
#endif
EXTERN_C const IID      IID_IShellRunDll;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IKnownFolder;
EXTERN_C const IID      IID_IKnownFolderManager;
EXTERN_C const IID      IID_ISharingConfigurationManager;
#endif
EXTERN_C const IID      IID_IPreviousVersionsInfo;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IRelatedItem;
EXTERN_C const IID      IID_IIdentityName;
EXTERN_C const IID      IID_IDelegateItem;
EXTERN_C const IID      IID_ICurrentItem;
EXTERN_C const IID      IID_ITransferMediumItem;
EXTERN_C const IID      IID_IUseToBrowseItem;
EXTERN_C const IID      IID_IDisplayItem;
EXTERN_C const IID      IID_IViewStateIdentityItem;
EXTERN_C const IID      IID_IPreviewItem;
#endif
EXTERN_C const IID      IID_IDestinationStreamFactory;
EXTERN_C const IID      IID_INewMenuClient;
#if (_WIN32_IE >= 0x0700)
EXTERN_C const IID      IID_IInitializeWithBindCtx;
EXTERN_C const IID      IID_IShellItemFilter;
#endif
EXTERN_C const IID      IID_INameSpaceTreeControl;
EXTERN_C const IID      IID_INameSpaceTreeControlEvents;
EXTERN_C const IID      IID_INameSpaceTreeAccessible;
EXTERN_C const IID      IID_INameSpaceTreeControlCustomDraw;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_INameSpaceTreeControlFolderCapabilities;
#endif
EXTERN_C const IID      IID_IPreviewHandler;
EXTERN_C const IID      IID_IPreviewHandlerFrame;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_ITrayDeskBand;
EXTERN_C const IID      IID_IBandHost;
EXTERN_C const IID      IID_IExplorerPaneVisibility;
EXTERN_C const IID      IID_IContextMenuCB;
#endif
EXTERN_C const IID      IID_IDefaultExtractIconInit;
EXTERN_C const IID      IID_IExplorerCommand;
EXTERN_C const IID      IID_IExplorerCommandState;
EXTERN_C const IID      IID_IInitializeCommand;
EXTERN_C const IID      IID_IEnumExplorerCommand;
EXTERN_C const IID      IID_IExplorerCommandProvider;
EXTERN_C const IID      IID_IMarkupCallback;
EXTERN_C const IID      IID_IControlMarkup;
EXTERN_C const IID      IID_IInitializeNetworkFolder;
EXTERN_C const IID      IID_IOpenControlPanel;
EXTERN_C const IID      IID_ISystemCPLUpdate;
EXTERN_C const IID      IID_IComputerInfoAdvise;
EXTERN_C const IID      IID_IComputerInfoChangeNotify;
EXTERN_C const IID      IID_IFileSystemBindData;
EXTERN_C const IID      IID_IFileSystemBindData2;
#if (NTDDI_VERSION >= 0x06010000)
EXTERN_C const IID      IID_ICustomDestinationList;
EXTERN_C const IID      IID_IApplicationDestinations;
EXTERN_C const IID      IID_IApplicationDocumentLists;
EXTERN_C const IID      IID_IObjectWithAppUserModelID;
EXTERN_C const IID      IID_IObjectWithProgID;
EXTERN_C const IID      IID_IUpdateIDList;
#endif
EXTERN_C const IID      IID_IDesktopGadget;
EXTERN_C const IID      IID_IHomeGroup;
EXTERN_C const IID      IID_IInitializeWithPropertyStore;
EXTERN_C const IID      IID_IOpenSearchSource;
EXTERN_C const IID      IID_IShellLibrary;
EXTERN_C const IID      LIBID_ShellObjects;
EXTERN_C const CLSID    CLSID_ShellDesktop;
EXTERN_C const CLSID    CLSID_ShellFSFolder;
EXTERN_C const CLSID    CLSID_NetworkPlaces;
EXTERN_C const CLSID    CLSID_ShellLink;
EXTERN_C const CLSID    CLSID_QueryCancelAutoPlay;
EXTERN_C const CLSID    CLSID_DriveSizeCategorizer;
EXTERN_C const CLSID    CLSID_DriveTypeCategorizer;
EXTERN_C const CLSID    CLSID_FreeSpaceCategorizer;
EXTERN_C const CLSID    CLSID_TimeCategorizer;
EXTERN_C const CLSID    CLSID_SizeCategorizer;
EXTERN_C const CLSID    CLSID_AlphabeticalCategorizer;
EXTERN_C const CLSID    CLSID_MergedCategorizer;
EXTERN_C const CLSID    CLSID_ImageProperties;
EXTERN_C const CLSID    CLSID_PropertiesUI;
EXTERN_C const CLSID    CLSID_UserNotification;
EXTERN_C const CLSID    CLSID_CDBurn;
EXTERN_C const CLSID    CLSID_TaskbarList;
EXTERN_C const CLSID    CLSID_StartMenuPin;
EXTERN_C const CLSID    CLSID_WebWizardHost;
EXTERN_C const CLSID    CLSID_PublishDropTarget;
EXTERN_C const CLSID    CLSID_PublishingWizard;
EXTERN_C const CLSID    CLSID_InternetPrintOrdering;
EXTERN_C const CLSID    CLSID_FolderViewHost;
EXTERN_C const CLSID    CLSID_ExplorerBrowser;
EXTERN_C const CLSID    CLSID_ImageRecompress;
EXTERN_C const CLSID    CLSID_TrayBandSiteService;
EXTERN_C const CLSID    CLSID_TrayDeskBand;
EXTERN_C const CLSID    CLSID_AttachmentServices;
EXTERN_C const CLSID    CLSID_DocPropShellExtension;
EXTERN_C const CLSID    CLSID_ShellItem;
EXTERN_C const CLSID    CLSID_NamespaceWalker;
EXTERN_C const CLSID    CLSID_FileOperation;
EXTERN_C const CLSID    CLSID_FileOpenDialog;
EXTERN_C const CLSID    CLSID_FileSaveDialog;
EXTERN_C const CLSID    CLSID_KnownFolderManager;
EXTERN_C const CLSID    CLSID_FSCopyHandler;
EXTERN_C const CLSID    CLSID_SharingConfigurationManager;
EXTERN_C const CLSID    CLSID_PreviousVersions;
EXTERN_C const CLSID    CLSID_NetworkConnections;
EXTERN_C const CLSID    CLSID_NamespaceTreeControl;
EXTERN_C const CLSID    CLSID_IENamespaceTreeControl;
EXTERN_C const CLSID    CLSID_ScheduledTasks;
EXTERN_C const CLSID    CLSID_ApplicationAssociationRegistration;
EXTERN_C const CLSID    CLSID_ApplicationAssociationRegistrationUI;
EXTERN_C const CLSID    CLSID_SearchFolderItemFactory;
EXTERN_C const CLSID    CLSID_OpenControlPanel;
EXTERN_C const CLSID    CLSID_ComputerInfoAdvise;
EXTERN_C const CLSID    CLSID_MailRecipient;
EXTERN_C const CLSID    CLSID_NetworkExplorerFolder;
EXTERN_C const CLSID    CLSID_DestinationList;
EXTERN_C const CLSID    CLSID_ApplicationDestinations;
EXTERN_C const CLSID    CLSID_ApplicationDocumentLists;
EXTERN_C const CLSID    CLSID_HomeGroup;
EXTERN_C const CLSID    CLSID_ShellLibrary;
EXTERN_C const CLSID    CLSID_AppStartupLink;
EXTERN_C const CLSID    CLSID_EnumerableObjectCollection;
EXTERN_C const CLSID    CLSID_DesktopGadget;
#if (NTDDI_VERSION >= 0x06000000)
EXTERN_C const IID      IID_IAssocHandlerInvoker;
EXTERN_C const IID      IID_IAssocHandler;
EXTERN_C const IID      IID_IEnumAssocHandler;
#endif

/* IContextMenu interface */
#undef INTERFACE
#define INTERFACE   IContextMenu
DECLARE_INTERFACE_( IContextMenu, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IContextMenu methods */
    STDMETHOD( QueryContextMenu )( THIS_ HMENU, UINT, UINT, UINT, UINT ) PURE;
    STDMETHOD( InvokeCommand )( THIS_ CMINVOKECOMMANDINFO * ) PURE;
    STDMETHOD( GetCommandString )( THIS_ UINT_PTR, UINT, UINT *, LPSTR, UINT ) PURE;
};
typedef IContextMenu    *LPCONTEXTMENU;

/* IContextMenu2 interface */
#undef INTERFACE
#define INTERFACE   IContextMenu2
DECLARE_INTERFACE_( IContextMenu2, IContextMenu ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IContextMenu methods */
    STDMETHOD( QueryContextMenu )( THIS_ HMENU, UINT, UINT, UINT, UINT ) PURE;
    STDMETHOD( InvokeCommand )( THIS_ CMINVOKECOMMANDINFO * ) PURE;
    STDMETHOD( GetCommandString )( THIS_ UINT_PTR, UINT, UINT *, LPSTR, UINT ) PURE;

    /* IContextMenu2 methods */
    STDMETHOD( HandleMenuMsg )( THIS_ UINT, WPARAM, LPARAM ) PURE;
};
typedef IContextMenu2   *LPCONTEXTMENU2;

/* IContextMenu3 interface */
#undef INTERFACE
#define INTERFACE   IContextMenu3
DECLARE_INTERFACE_( IContextMenu3, IContextMenu2 ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IContextMenu methods */
    STDMETHOD( QueryContextMenu )( THIS_ HMENU, UINT, UINT, UINT, UINT ) PURE;
    STDMETHOD( InvokeCommand )( THIS_ CMINVOKECOMMANDINFO * ) PURE;
    STDMETHOD( GetCommandString )( THIS_ UINT_PTR, UINT, UINT *, LPSTR, UINT ) PURE;

    /* IContextMenu2 methods */
    STDMETHOD( HandleMenuMsg )( THIS_ UINT, WPARAM, LPARAM ) PURE;

    /* IContextMenu3 methods */
    STDMETHOD( HandleMenuMsg2 )( THIS_ UINT, WPARAM, LPARAM, LRESULT * ) PURE;
};
typedef IContextMenu3   *LPCONTEXTMENU3;

/* IExecuteCommand interface */
#undef INTERFACE
#define INTERFACE   IExecuteCommand
DECLARE_INTERFACE_( IExecuteCommand, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExecuteCommand methods */
    STDMETHOD( SetKeyState )( THIS_ DWORD ) PURE;
    STDMETHOD( SetParameters )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetPosition )( THIS_ POINT ) PURE;
    STDMETHOD( SetShowWindow )( THIS_ int ) PURE;
    STDMETHOD( SetNoShowUI )( THIS_ BOOL ) PURE;
    STDMETHOD( SetDirectory )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Execute )( THIS ) PURE;
};

/* IPersistFolder interface */
#undef INTERFACE
#define INTERFACE   IPersistFolder
DECLARE_INTERFACE_( IPersistFolder, IPersist ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IPersist methods */
    STDMETHOD( GetClassID )( THIS_ CLSID * ) PURE;

    /* IPersistFolder methods */
    STDMETHOD( Initialize )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
};
typedef IPersistFolder  *LPPERSISTFOLDER;

/* IRunnableTask interface */
#if (_WIN32_IE >= 0x0400)
#undef INTERFACE
#define INTERFACE   IRunnableTask
DECLARE_INTERFACE_( IRunnableTask, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRunnableTask methods */
    STDMETHOD( Run )( THIS ) PURE;
    STDMETHOD( Kill )( THIS_ BOOL ) PURE;
    STDMETHOD( Suspend )( THIS ) PURE;
    STDMETHOD( Resume )( THIS ) PURE;
    STDMETHOD_( ULONG, IsRunning )( THIS ) PURE;
};
#endif

/* IShellTaskScheduler interface */
#if (_WIN32_IE >= 0x0400)
#undef INTERFACE
#define INTERFACE   IShellTaskScheduler
DECLARE_INTERFACE_( IShellTaskScheduler, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellTaskScheduler methods */
    STDMETHOD( AddTask )( THIS_ IRunnableTask *, REFTASKOWNERID, DWORD_PTR, DWORD ) PURE;
    STDMETHOD( RemoveTasks )( THIS_ REFTASKOWNERID, DWORD_PTR, BOOL ) PURE;
    STDMETHOD_( UINT, CountTasks )( THIS_ REFTASKOWNERID ) PURE;
    STDMETHOD( Status )( THIS_ DWORD, DWORD ) PURE;
};
#endif

/* IQueryCodePage interface */
#if (_WIN32_IE >= 0x0400)
#undef INTERFACE
#define INTERFACE   IQueryCodePage
DECLARE_INTERFACE_( IQueryCodePage, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IQueryCodePage methods */
    STDMETHOD( GetCodePage )( THIS_ UINT * ) PURE;
    STDMETHOD( SetCodePage )( THIS_ UINT ) PURE;
};
#endif

/* IPersistFolder2 interface */
#if (_WIN32_IE >= 0x0400)
#undef INTERFACE
#define INTERFACE   IPersistFolder2
DECLARE_INTERFACE_( IPersistFolder2, IPersistFolder ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IPersist methods */
    STDMETHOD( GetClassID )( THIS_ CLSID * ) PURE;

    /* IPersistFolder methods */
    STDMETHOD( Initialize )( THIS_ PCIDLIST_ABSOLUTE ) PURE;

    /* IPersistFolder2 methods */
    STDMETHOD( GetCurFolder )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
};
#endif

/* IPersistFolder3 interface */
#if (_WIN32_IE >= 0x0500)
#undef INTERFACE
#define INTERFACE   IPersistFolder3
DECLARE_INTERFACE_( IPersistFolder3, IPersistFolder2 ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IPersist methods */
    STDMETHOD( GetClassID )( THIS_ CLSID * ) PURE;

    /* IPersistFolder methods */
    STDMETHOD( Initialize )( THIS_ PCIDLIST_ABSOLUTE ) PURE;

    /* IPersistFolder2 methods */
    STDMETHOD( GetCurFolder )( THIS_ PIDLIST_ABSOLUTE * ) PURE;

    /* IPersistFolder3 methods */
    STDMETHOD( InitializeEx )( THIS_ IBindCtx *, PCIDLIST_ABSOLUTE, const PERSIST_FOLDER_TARGET_INFO * ) PURE;
    STDMETHOD( GetFolderTargetInfo )( THIS_ PERSIST_FOLDER_TARGET_INFO * ) PURE;
};
#endif

/* IPersistIDList interface */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0600)
#undef INTERFACE
#define INTERFACE   IPersistIDList
DECLARE_INTERFACE_( IPersistIDList, IPersist ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IPersist methods */
    STDMETHOD( GetClassID )( THIS_ CLSID * ) PURE;

    /* IPersistIDList methods */
    STDMETHOD( SetIDList )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
    STDMETHOD( GetIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
};
#endif

/* IEnumIDList interface */
#undef INTERFACE
#define INTERFACE   IEnumIDList
DECLARE_INTERFACE_( IEnumIDList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumIDList methods */
    STDMETHOD( Next )( THIS_ ULONG, PITEMID_CHILD *, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumIDList ** ) PURE;
};
typedef IEnumIDList *LPENUMIDLIST;

/* IEnumFullIDList interface */
#undef INTERFACE
#define INTERFACE   IEnumFullIDList
DECLARE_INTERFACE_( IEnumFullIDList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumFullIDList methods */
    STDMETHOD( Next )( THIS_ ULONG, PIDLIST_ABSOLUTE *, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumFullIDList ** ) PURE;
};

/* IObjectWithFolderEnumMode interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IObjectWithFolderEnumMode
DECLARE_INTERFACE_( IObjectWithFolderEnumMode, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectWithFolderEnumMode interface */
    STDMETHOD( SetMode )( THIS_ FOLDER_ENUM_MODE );
    STDMETHOD( GetMode )( THIS_ FOLDER_ENUM_MODE * );
};
#endif

/* IParseAndCreateItem interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IParseAndCreateItem
DECLARE_INTERFACE_( IParseAndCreateItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IParseAndCreateItem methods */
    STDMETHOD( SetItem )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetItem )( THIS_ REFIID, void ** ) PURE;
};
#endif

/* IShellFolder interface */
#undef INTERFACE
#define INTERFACE   IShellFolder
DECLARE_INTERFACE_( IShellFolder, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellFolder methods */
    STDMETHOD( ParseDisplayName )( THIS_ HWND, IBindCtx *, LPWSTR, ULONG *, PIDLIST_RELATIVE *, ULONG * ) PURE;
    STDMETHOD( EnumObjects )( THIS_ HWND, SHCONTF, IEnumIDList ** ) PURE;
    STDMETHOD( BindToObject )( THIS_ PCUIDLIST_RELATIVE, IBindCtx *, REFIID, void ** ) PURE;
    STDMETHOD( BindToStorage )( THIS_ PCUIDLIST_RELATIVE, IBindCtx *, REFIID, void ** ) PURE;
    STDMETHOD( CompareIDs )( THIS_ LPARAM, PCUIDLIST_RELATIVE, PCUIDLIST_RELATIVE ) PURE;
    STDMETHOD( CreateViewObject )( THIS_ HWND, REFIID, void ** ) PURE;
    STDMETHOD( GetAttributesOf )( THIS_ UINT, PCUITEMID_CHILD_ARRAY, SFGAOF * ) PURE;
    STDMETHOD( GetUIObjectOf )( THIS_ HWND, UINT, PCUITEMID_CHILD_ARRAY, REFIID, UINT *, void ** ) PURE;
    STDMETHOD( GetDisplayNameOf )( THIS_ PCUITEMID_CHILD, SHGDNF, STRRET * ) PURE;
    STDMETHOD( SetNameOf )( THIS_ HWND, PCUITEMID_CHILD, LPCWSTR, SHGDNF, PITEMID_CHILD * ) PURE;
};
typedef IShellFolder    *LPSHELLFOLDER;

/* IEnumExtraSearch interface */
#undef INTERFACE
#define INTERFACE   IEnumExtraSearch
DECLARE_INTERFACE_( IEnumExtraSearch, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumExtraSearch methods */
    STDMETHOD( Next )( THIS_ ULONG, EXTRASEARCH *, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumExtraSearch ** ) PURE;
};
typedef IEnumExtraSearch    *LPENUMEXTRASEARCH;

/* IShellFolder2 interface */
#undef INTERFACE
#define INTERFACE   IShellFolder2
DECLARE_INTERFACE_( IShellFolder2, IShellFolder ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellFolder methods */
    STDMETHOD( ParseDisplayName )( THIS_ HWND, IBindCtx *, LPWSTR, ULONG *, PIDLIST_RELATIVE *, ULONG * ) PURE;
    STDMETHOD( EnumObjects )( THIS_ HWND, SHCONTF, IEnumIDList ** ) PURE;
    STDMETHOD( BindToObject )( THIS_ PCUIDLIST_RELATIVE, IBindCtx *, REFIID, void ** ) PURE;
    STDMETHOD( BindToStorage )( THIS_ PCUIDLIST_RELATIVE, IBindCtx *, REFIID, void ** ) PURE;
    STDMETHOD( CompareIDs )( THIS_ LPARAM, PCUIDLIST_RELATIVE, PCUIDLIST_RELATIVE ) PURE;
    STDMETHOD( CreateViewObject )( THIS_ HWND, REFIID, void ** ) PURE;
    STDMETHOD( GetAttributesOf )( THIS_ UINT, PCUITEMID_CHILD_ARRAY, SFGAOF * ) PURE;
    STDMETHOD( GetUIObjectOf )( THIS_ HWND, UINT, PCUITEMID_CHILD_ARRAY, REFIID, UINT *, void ** ) PURE;
    STDMETHOD( GetDisplayNameOf )( THIS_ PCUITEMID_CHILD, SHGDNF, STRRET * ) PURE;
    STDMETHOD( SetNameOf )( THIS_ HWND, PCUITEMID_CHILD, LPCWSTR, SHGDNF, PITEMID_CHILD * ) PURE;

    /* IShellFolder2 methods */
    STDMETHOD( GetDefaultSearchGUID )( THIS_ GUID * ) PURE;
    STDMETHOD( EnumSearches )( THIS_ IEnumExtraSearch ** ) PURE;
    STDMETHOD( GetDefaultColumn )( THIS_ DWORD, ULONG *, ULONG * ) PURE;
    STDMETHOD( GetDefaultColumnState )( THIS_ UINT, SHCOLSTATEF * ) PURE;
    STDMETHOD( GetDetailsEx )( THIS_ PCUITEMID_CHILD, const SHCOLUMNID *, VARIANT * ) PURE;
    STDMETHOD( GetDetailsOf )( THIS_ PCUITEMID_CHILD, UINT, SHELLDETAILS * ) PURE;
    STDMETHOD( MapColumnToSCID )( THIS_ UINT, SHCOLUMNID * ) PURE;
};

/* IFolderViewOptions interface */
#undef INTERFACE
#define INTERFACE   IFolderViewOptions
DECLARE_INTERFACE_( IFolderViewOptions, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderViewOptions methods */
    STDMETHOD( SetFolderViewOptions )( THIS_ FOLDERVIEWOPTIONS, FOLDERVIEWOPTIONS ) PURE;
    STDMETHOD( GetFolderViewOptions )( THIS_ FOLDERVIEWOPTIONS * ) PURE;
};

/* IShellView interface */
#undef INTERFACE
#define INTERFACE   IShellView
DECLARE_INTERFACE_( IShellView, IOleWindow ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IShellView methods */
    STDMETHOD( TranslateAccelerator )( THIS_ MSG * ) PURE;
    STDMETHOD( EnableModeless )( THIS_ BOOL ) PURE;
    STDMETHOD( UIActivate )( THIS_ UINT ) PURE;
    STDMETHOD( Refresh )( THIS ) PURE;
    STDMETHOD( CreateViewWindow )( THIS_ IShellView *, LPCFOLDERSETTINGS, IShellBrowser *, RECT *, HWND * ) PURE;
    STDMETHOD( DestroyViewWindow )( THIS ) PURE;
    STDMETHOD( GetCurrentInfo )( THIS_ LPFOLDERSETTINGS ) PURE;
    STDMETHOD( AddPropertySheetPages )( THIS_ DWORD, LPFNSVADDPROPSHEETPAGE, LPARAM ) PURE;
    STDMETHOD( SaveViewState )( THIS ) PURE;
    STDMETHOD( SelectItem )( THIS_ PCUITEMID_CHILD, SVSIF ) PURE;
    STDMETHOD( GetItemObject )( THIS_ UINT, REFIID, void ** ) PURE;
};
typedef IShellView  *LPSHELLVIEW;

/* IShellView2 interface */
#undef INTERFACE
#define INTERFACE   IShellView2
DECLARE_INTERFACE_( IShellView2, IShellView ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IShellView methods */
    STDMETHOD( TranslateAccelerator )( THIS_ MSG * ) PURE;
    STDMETHOD( EnableModeless )( THIS_ BOOL ) PURE;
    STDMETHOD( UIActivate )( THIS_ UINT ) PURE;
    STDMETHOD( Refresh )( THIS ) PURE;
    STDMETHOD( CreateViewWindow )( THIS_ IShellView *, LPCFOLDERSETTINGS, IShellBrowser *, RECT *, HWND * ) PURE;
    STDMETHOD( DestroyViewWindow )( THIS ) PURE;
    STDMETHOD( GetCurrentInfo )( THIS_ LPFOLDERSETTINGS ) PURE;
    STDMETHOD( AddPropertySheetPages )( THIS_ DWORD, LPFNSVADDPROPSHEETPAGE, LPARAM ) PURE;
    STDMETHOD( SaveViewState )( THIS ) PURE;
    STDMETHOD( SelectItem )( THIS_ PCUITEMID_CHILD, SVSIF ) PURE;
    STDMETHOD( GetItemObject )( THIS_ UINT, REFIID, void ** ) PURE;

    /* IShellView2 methods */
    STDMETHOD( GetView )( THIS_ SHELLVIEWID *, ULONG ) PURE;
    STDMETHOD( CreateViewWindow2 )( THIS_ LPSV2CVW2_PARAMS ) PURE;
    STDMETHOD( HandleResume )( THIS_ PCUITEMID_CHILD ) PURE;
    STDMETHOD( SelectAndPositionItem )( THIS_ PCUITEMID_CHILD, UINT, POINT * ) PURE;
};

/* IShellView3 interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IShellView3
DECLARE_INTERFACE_( IShellView3, IShellView2 ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IShellView methods */
    STDMETHOD( TranslateAccelerator )( THIS_ MSG * ) PURE;
    STDMETHOD( EnableModeless )( THIS_ BOOL ) PURE;
    STDMETHOD( UIActivate )( THIS_ UINT ) PURE;
    STDMETHOD( Refresh )( THIS ) PURE;
    STDMETHOD( CreateViewWindow )( THIS_ IShellView *, LPCFOLDERSETTINGS, IShellBrowser *, RECT *, HWND * ) PURE;
    STDMETHOD( DestroyViewWindow )( THIS ) PURE;
    STDMETHOD( GetCurrentInfo )( THIS_ LPFOLDERSETTINGS ) PURE;
    STDMETHOD( AddPropertySheetPages )( THIS_ DWORD, LPFNSVADDPROPSHEETPAGE, LPARAM ) PURE;
    STDMETHOD( SaveViewState )( THIS ) PURE;
    STDMETHOD( SelectItem )( THIS_ PCUITEMID_CHILD, SVSIF ) PURE;
    STDMETHOD( GetItemObject )( THIS_ UINT, REFIID, void ** ) PURE;

    /* IShellView2 methods */
    STDMETHOD( GetView )( THIS_ SHELLVIEWID *, ULONG ) PURE;
    STDMETHOD( CreateViewWindow2 )( THIS_ LPSV2CVW2_PARAMS ) PURE;
    STDMETHOD( HandleResume )( THIS_ PCUITEMID_CHILD ) PURE;
    STDMETHOD( SelectAndPositionItem )( THIS_ PCUITEMID_CHILD, UINT, POINT * ) PURE;

    /* IShellView3 methods */
    STDMETHOD( CreateViewWindow3 )( THIS_ IShellBrowser *, IShellView *, SV3CVW3_FLAGS, FOLDERFLAGS, FOLDERFLAGS, FOLDERVIEWMODE, const SHELLVIEWID *, const RECT *, HWND * ) PURE;
};
#endif

/* IFolderView interface */
#undef INTERFACE
#define INTERFACE   IFolderView
DECLARE_INTERFACE_( IFolderView, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderView methods */
    STDMETHOD( GetCurrentViewMode )( THIS_ UINT * ) PURE;
    STDMETHOD( SetCurrentViewMode )( THIS_ UINT ) PURE;
    STDMETHOD( GetFolder )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( Item )( THIS_ int, PITEMID_CHILD * ) PURE;
    STDMETHOD( ItemCount )( THIS_ UINT, int * ) PURE;
    STDMETHOD( Items )( THIS_ UINT, REFIID, void ** ) PURE;
    STDMETHOD( GetSelectionMarkedItem )( THIS_ int * ) PURE;
    STDMETHOD( GetFocusedItem )( THIS_ int * ) PURE;
    STDMETHOD( GetItemPosition )( THIS_ PCUITEMID_CHILD, POINT * ) PURE;
    STDMETHOD( GetSpacing )( THIS_ POINT * ) PURE;
    STDMETHOD( GetDefaultSpacing )( THIS_ POINT * ) PURE;
    STDMETHOD( GetAutoArrange )( THIS ) PURE;
    STDMETHOD( SelectItem )( THIS_ int, DWORD ) PURE;
    STDMETHOD( SelectAndPositionItems )( THIS_ UINT, PCUITEMID_CHILD_ARRAY, POINT *, DWORD ) PURE;
};

/* ISearchBoxInfo interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   ISearchBoxInfo
DECLARE_INTERFACE_( ISearchBoxInfo, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ISearchBoxInfo methods */
    STDMETHOD( GetCondition )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( GetText )( THIS_ LPWSTR * ) PURE;
};
#endif

/* IFolderView2 interface */
#if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IFolderView2
DECLARE_INTERFACE_( IFolderView2, IFolderView ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderView methods */
    STDMETHOD( GetCurrentViewMode )( THIS_ UINT * ) PURE;
    STDMETHOD( SetCurrentViewMode )( THIS_ UINT ) PURE;
    STDMETHOD( GetFolder )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( Item )( THIS_ int, PITEMID_CHILD * ) PURE;
    STDMETHOD( ItemCount )( THIS_ UINT, int * ) PURE;
    STDMETHOD( Items )( THIS_ UINT, REFIID, void ** ) PURE;
    STDMETHOD( GetSelectionMarkedItem )( THIS_ int * ) PURE;
    STDMETHOD( GetFocusedItem )( THIS_ int * ) PURE;
    STDMETHOD( GetItemPosition )( THIS_ PCUITEMID_CHILD, POINT * ) PURE;
    STDMETHOD( GetSpacing )( THIS_ POINT * ) PURE;
    STDMETHOD( GetDefaultSpacing )( THIS_ POINT * ) PURE;
    STDMETHOD( GetAutoArrange )( THIS ) PURE;
    STDMETHOD( SelectItem )( THIS_ int, DWORD ) PURE;
    STDMETHOD( SelectAndPositionItems )( THIS_ UINT, PCUITEMID_CHILD_ARRAY, POINT *, DWORD ) PURE;

    /* IFolderView2 methods */
    STDMETHOD( SetGroupBy )( THIS_ REFPROPERTYKEY, BOOL ) PURE;
    STDMETHOD( GetGroupBy )( THIS_ PROPERTYKEY *, BOOL * ) PURE;
    STDMETHOD( SetViewProperty )( THIS_ PCUITEMID_CHILD, REFPROPERTYKEY, REFPROPVARIANT ) PURE;
    STDMETHOD( GetViewProperty )( THIS_ PCUITEMID_CHILD, REFPROPERTYKEY, PROPVARIANT * ) PURE;
    STDMETHOD( SetTileViewProperties )( THIS_ PCUITEMID_CHILD, LPCWSTR ) PURE;
    STDMETHOD( SetExtendedTileViewProperties )( THIS_ PCUITEMID_CHILD, LPCWSTR ) PURE;
    STDMETHOD( SetText )( THIS_ FVTEXTTYPE, LPCWSTR ) PURE;
    STDMETHOD( SetCurrentFolderFlags )( THIS_ DWORD, DWORD ) PURE;
    STDMETHOD( GetCurrentFolderFlags )( THIS_ DWORD * ) PURE;
    STDMETHOD( GetSortColumnCount )( THIS_ int * ) PURE;
    STDMETHOD( SetSortColumns )( THIS_ const SORTCOLUMN *, int ) PURE;
    STDMETHOD( GetSortColumns )( THIS_ SORTCOLUMN *, int ) PURE;
    STDMETHOD( GetItem )( THIS_ int, REFIID, void ** ) PURE;
    STDMETHOD( GetVisibleItem )( THIS_ int, BOOL, int * ) PURE;
    STDMETHOD( GetSelectedItem )( THIS_ int, int * ) PURE;
    STDMETHOD( GetSelection )( THIS_ BOOL, IShellItemArray ** ) PURE;
    STDMETHOD( GetSelectionState )( THIS_ PCUITEMID_CHILD, DWORD * ) PURE;
    STDMETHOD( InvokeVerbOnSelection )( THIS_ LPCSTR ) PURE;
    STDMETHOD( SetViewModeAndIconSize )( THIS_ FOLDERVIEWMODE, int ) PURE;
    STDMETHOD( GetViewModeAndIconSize )( THIS_ FOLDERVIEWMODE *, int * ) PURE;
    STDMETHOD( SetGroupSubsetCount )( THIS_ UINT ) PURE;
    STDMETHOD( GetGroupSubsetCount )( THIS_ UINT * ) PURE;
    STDMETHOD( SetRedraw )( THIS_ BOOL ) PURE;
    STDMETHOD( IsMoveInSameFolder )( THIS ) PURE;
    STDMETHOD( DoRename )( THIS ) PURE;
};
#endif

/* IFolderViewSettings interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFolderViewSettings
DECLARE_INTERFACE_( IFolderViewSettings, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderViewSettings methods */
    STDMETHOD( GetColumnPropertyList )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( GetGroupByProperty )( THIS_ PROPERTYKEY *, BOOL * ) PURE;
    STDMETHOD( GetViewMode )( THIS_ FOLDERLOGICALVIEWMODE * ) PURE;
    STDMETHOD( GetIconSize )( THIS_ UINT * ) PURE;
    STDMETHOD( GetFolderFlags )( THIS_ FOLDERFLAGS *, FOLDERFLAGS * ) PURE;
    STDMETHOD( GetSortColumns )( THIS_ SORTCOLUMN *, UINT, UINT * ) PURE;
    STDMETHOD( GetGroupSubsetCount )( THIS_ UINT * ) PURE;
};
#endif

/* IPreviewHandlerVisuals interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IPreviewHandlerVisuals
DECLARE_INTERFACE_( IPreviewHandlerVisuals, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IPreviewHandlerVisuals methods */
    STDMETHOD( SetBackgroundColor )( THIS_ COLORREF ) PURE;
    STDMETHOD( SetFont )( THIS_ const LOGFONTW * ) PURE;
    STDMETHOD( SetTextColor )( THIS_ COLORREF ) PURE;
};
#endif

/* IVisualProperties interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IVisualProperties
DECLARE_INTERFACE_( IVisualProperties, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IVisualProperties methods */
    STDMETHOD( SetWatermark )( THIS_ HBITMAP, VPWATERMARKFLAGS ) PURE;
    STDMETHOD( SetColor )( THIS_ VPCOLORFLAGS, COLORREF ) PURE;
    STDMETHOD( GetColor )( THIS_ VPCOLORFLAGS, COLORREF * ) PURE;
    STDMETHOD( SetItemHeight )( THIS_ int ) PURE;
    STDMETHOD( GetItemHeight )( THIS_ int * ) PURE;
    STDMETHOD( SetFont )( THIS_ const LOGFONTW *, BOOL ) PURE;
    STDMETHOD( GetFont )( THIS_ LOGFONTW * ) PURE;
    STDMETHOD( SetTheme )( THIS_ LPCWSTR, LPCWSTR ) PURE;
};
#endif

/* ICommDlgBrowser interface */
#undef INTERFACE
#define INTERFACE   ICommDlgBrowser
DECLARE_INTERFACE_( ICommDlgBrowser, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICommDlgBrowser methods */
    STDMETHOD( OnDefaultCommand )( THIS_ IShellView * ) PURE;
    STDMETHOD( OnStateChange )( THIS_ IShellView *, ULONG ) PURE;
    STDMETHOD( IncludeObject )( THIS_ IShellView *, PCUITEMID_CHILD ) PURE;
};
typedef ICommDlgBrowser *LPCOMMDLGBROWSER;

/* ICommDlgBrowser2 interface */
#if (NTDDI_VERSION >= 0x05000000)
#undef INTERFACE
#define INTERFACE   ICommDlgBrowser2
DECLARE_INTERFACE_( ICommDlgBrowser2, ICommDlgBrowser ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICommDlgBrowser methods */
    STDMETHOD( OnDefaultCommand )( THIS_ IShellView * ) PURE;
    STDMETHOD( OnStateChange )( THIS_ IShellView *, ULONG ) PURE;
    STDMETHOD( IncludeObject )( THIS_ IShellView *, PCUITEMID_CHILD ) PURE;

    /* ICommDlgBrowser2 methods */
    STDMETHOD( Notify )( THIS_ IShellView *, DWORD ) PURE;
    STDMETHOD( GetDefaultMenuText )( THIS_ IShellView *, LPWSTR, int ) PURE;
    STDMETHOD( GetViewFlags )( THIS_ DWORD * ) PURE;
};
typedef ICommDlgBrowser2    *LPCOMMDLGBROWSER2;
#endif

/* ICommDlgBrowser3 interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   ICommDlgBrowser3
DECLARE_INTERFACE_( ICommDlgBrowser3, ICommDlgBrowser2 ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICommDlgBrowser methods */
    STDMETHOD( OnDefaultCommand )( THIS_ IShellView * ) PURE;
    STDMETHOD( OnStateChange )( THIS_ IShellView *, ULONG ) PURE;
    STDMETHOD( IncludeObject )( THIS_ IShellView *, PCUITEMID_CHILD ) PURE;

    /* ICommDlgBrowser2 methods */
    STDMETHOD( Notify )( THIS_ IShellView *, DWORD ) PURE;
    STDMETHOD( GetDefaultMenuText )( THIS_ IShellView *, LPWSTR, int ) PURE;
    STDMETHOD( GetViewFlags )( THIS_ DWORD * ) PURE;

    /* ICommDlgBrowser3 methods */
    STDMETHOD( OnColumnClicked )( THIS_ IShellView *, int ) PURE;
    STDMETHOD( GetCurrentFilter )( THIS_ LPWSTR, int ) PURE;
    STDMETHOD( OnPreViewCreated )( THIS_ IShellView * ) PURE;
};
#endif

/* IColumnManager interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IColumnManager
DECLARE_INTERFACE_( IColumnManager, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IColumnManager methods */
    STDMETHOD( SetColumnInfo )( THIS_ REFPROPERTYKEY, const CM_COLUMNINFO * ) PURE;
    STDMETHOD( GetColumnInfo )( THIS_ REFPROPERTYKEY, CM_COLUMNINFO * ) PURE;
    STDMETHOD( GetColumnCount )( THIS_ CM_ENUM_FLAGS, UINT * ) PURE;
    STDMETHOD( GetColumns )( THIS_ CM_ENUM_FLAGS, PROPERTYKEY *, UINT ) PURE;
    STDMETHOD( SetColumns )( THIS_ const PROPERTYKEY *, UINT ) PURE;
};
#endif

/* IFolderFilterSite interface */
#undef INTERFACE
#define INTERFACE   IFolderFilterSite
DECLARE_INTERFACE_( IFolderFilterSite, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderFilterSite methods */
    STDMETHOD( SetFilter )( THIS_ IUnknown * ) PURE;
};

/* IFolderFilter interface */
#undef INTERFACE
#define INTERFACE   IFolderFilter
DECLARE_INTERFACE_( IFolderFilter, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderFilter methods */
    STDMETHOD( ShouldShow )( THIS_ IShellFolder *, PCIDLIST_ABSOLUTE, PCUITEMID_CHILD ) PURE;
    STDMETHOD( GetEnumFlags )( THIS_ IShellFolder *, PCIDLIST_ABSOLUTE, HWND *, DWORD * ) PURE;
};

/* IInputObjectSite interface */
#undef INTERFACE
#define INTERFACE   IInputObjectSite
DECLARE_INTERFACE_( IInputObjectSite, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInputObjectSite methods */
    STDMETHOD( OnFocusChangeIS )( THIS_ IUnknown *, BOOL ) PURE;
};

/* IInputObject interface */
#undef INTERFACE
#define INTERFACE   IInputObject
DECLARE_INTERFACE_( IInputObject, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInputObject methods */
    STDMETHOD( UIActivateIO )( THIS_ BOOL, MSG * ) PURE;
    STDMETHOD( HasFocusIO )( THIS ) PURE;
    STDMETHOD( TranslateAcceleratorIO )( THIS_ MSG * ) PURE;
};

/* IInputObject2 interface */
#undef INTERFACE
#define INTERFACE   IInputObject2
DECLARE_INTERFACE_( IInputObject2, IInputObject ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInputObject methods */
    STDMETHOD( UIActivateIO )( THIS_ BOOL, MSG * ) PURE;
    STDMETHOD( HasFocusIO )( THIS ) PURE;
    STDMETHOD( TranslateAcceleratorIO )( THIS_ MSG * ) PURE;

    /* IInputObject2 methods */
    STDMETHOD( TranslateAcceleratorGlobal )( THIS_ MSG * ) PURE;
};

/* IShellIcon interface */
#undef INTERFACE
#define INTERFACE   IShellIcon
DECLARE_INTERFACE_( IShellIcon, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellIcon methods */
    STDMETHOD( GetIconOf )( THIS_ PCUITEMID_CHILD, UINT, int * ) PURE;
};

/* IShellBrowser interface */
#undef INTERFACE
#define INTERFACE   IShellBrowser
DECLARE_INTERFACE_( IShellBrowser, IOleWindow ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IShellBrowser methods */
    STDMETHOD( InsertMenusSB )( THIS_ HMENU, LPOLEMENUGROUPWIDTHS ) PURE;
    STDMETHOD( SetMenuSB )( THIS_ HMENU, HOLEMENU, HWND ) PURE;
    STDMETHOD( RemoveMenusSB )( THIS_ HMENU ) PURE;
    STDMETHOD( SetStatusTextSB )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( EnableModelessSB )( THIS_ BOOL ) PURE;
    STDMETHOD( TranslateAcceleratorSB )( THIS_ MSG *, WORD ) PURE;
    STDMETHOD( BrowseObject )( THIS_ PCUIDLIST_RELATIVE, UINT ) PURE;
    STDMETHOD( GetViewStateStream )( THIS_ DWORD, IStream ** ) PURE;
    STDMETHOD( GetControlWindow )( THIS_ UINT, HWND * ) PURE;
    STDMETHOD( SendControlMsg )( THIS_ UINT, UINT, WPARAM, LPARAM, LRESULT * ) PURE;
    STDMETHOD( QueryActiveShellView )( THIS_ IShellView ** ) PURE;
    STDMETHOD( OnViewWindowActive )( THIS_ IShellView * ) PURE;
    STDMETHOD( SetToolbarItems )( THIS_ LPTBBUTTONSB, UINT, UINT ) PURE;
};
typedef IShellBrowser   *LPSHELLBROWSER;

/* IProfferService interface */
#undef INTERFACE
#define INTERFACE   IProfferService
DECLARE_INTERFACE_( IProfferService, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IProfferService methods */
    STDMETHOD( ProfferService )( THIS_ REFGUID, IServiceProvider *, DWORD * ) PURE;
    STDMETHOD( RevokeService )( THIS_ DWORD ) PURE;
};

/* IShellItem interface */
#undef INTERFACE
#define INTERFACE   IShellItem
DECLARE_INTERFACE_( IShellItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellItem methods */
    STDMETHOD( BindToHandler )( THIS_ IBindCtx *, REFGUID, REFIID, void ** ) PURE;
    STDMETHOD( GetParent )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( GetDisplayName )( THIS_ SIGDN, LPWSTR * ) PURE;
    STDMETHOD( GetAttributes )( THIS_ SFGAOF, SFGAOF * ) PURE;
    STDMETHOD( Compare )( THIS_ IShellItem *, SICHINTF, int * ) PURE;
};

/* IShellItem2 interface */
#undef INTERFACE
#define INTERFACE   IShellItem2
DECLARE_INTERFACE_( IShellItem2, IShellItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellItem methods */
    STDMETHOD( BindToHandler )( THIS_ IBindCtx *, REFGUID, REFIID, void ** ) PURE;
    STDMETHOD( GetParent )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( GetDisplayName )( THIS_ SIGDN, LPWSTR * ) PURE;
    STDMETHOD( GetAttributes )( THIS_ SFGAOF, SFGAOF * ) PURE;
    STDMETHOD( Compare )( THIS_ IShellItem *, SICHINTF, int * ) PURE;

    /* IShellItem2 methods */
    STDMETHOD( GetPropertyStore )( THIS_ GETPROPERTYSTOREFLAGS, REFIID, void ** ) PURE;
    STDMETHOD( GetPropertyStoreWithCreateObject )( THIS_ GETPROPERTYSTOREFLAGS, IUnknown *, REFIID, void ** ) PURE;
    STDMETHOD( GetPropertyStoreForKeys )( THIS_ const PROPERTYKEY *, UINT, GETPROPERTYSTOREFLAGS, REFIID, void ** ) PURE;
    STDMETHOD( GetPropertyDescriptionList )( THIS_ REFPROPERTYKEY, REFIID, void ** ) PURE;
    STDMETHOD( Update )( THIS_ IBindCtx * ) PURE;
    STDMETHOD( GetProperty )( THIS_ REFPROPERTYKEY, PROPVARIANT * ) PURE;
    STDMETHOD( GetCLSID )( THIS_ REFPROPERTYKEY, CLSID * ) PURE;
    STDMETHOD( GetFileTime )( THIS_ REFPROPERTYKEY, FILETIME * ) PURE;
    STDMETHOD( GetInt32 )( THIS_ REFPROPERTYKEY, int * ) PURE;
    STDMETHOD( GetString )( THIS_ REFPROPERTYKEY, LPWSTR * ) PURE;
    STDMETHOD( GetUInt32 )( THIS_ REFPROPERTYKEY, ULONG * ) PURE;
    STDMETHOD( GetUInt64 )( THIS_ REFPROPERTYKEY, ULONGLONG * ) PURE;
    STDMETHOD( GetBool )( THIS_ REFPROPERTYKEY, BOOL * ) PURE;
};

/* IShellItemImageFactory interface */
#undef INTERFACE
#define INTERFACE   IShellItemImageFactory
DECLARE_INTERFACE_( IShellItemImageFactory, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellItemImageFactory methods */
    STDMETHOD( GetImage )( THIS_ SIZE, SIIGBF, HBITMAP * ) PURE;
};

/* IUserAccountChangeCallback interface */
#undef INTERFACE
#define INTERFACE   IUserAccountChangeCallback
DECLARE_INTERFACE_( IUserAccountChangeCallback, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IUserAccountChangeCallback methods */
    STDMETHOD( OnPictureChange )( THIS_ LPCWSTR ) PURE;
};

/* IEnumShellItems interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   IEnumShellItems
DECLARE_INTERFACE_( IEnumShellItems, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumShellItems methods */
    STDMETHOD( Next )( THIS_ ULONG, IShellItem **, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumShellItems ** ) PURE;
};
#endif

/* ITransferAdviseSink interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   ITransferAdviseSink
DECLARE_INTERFACE_( ITransferAdviseSink, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITransferAdviseSink methods */
    STDMETHOD( UpdateProgress )( THIS_ ULONGLONG, ULONGLONG, int, int, int, int ) PURE;
    STDMETHOD( UpdateTransferState )( THIS_ TRANSFER_ADVISE_STATE ) PURE;
    STDMETHOD( ConfirmOverwrite )( THIS_ IShellItem *, IShellItem *, LPCWSTR ) PURE;
    STDMETHOD( ConfirmEncryptionLoss )( THIS_ IShellItem * ) PURE;
    STDMETHOD( FileFailure )( THIS_ IShellItem *, LPCWSTR, HRESULT, LPWSTR, ULONG ) PURE;
    STDMETHOD( SubStreamFailure )( THIS_ IShellItem *, LPCWSTR, HRESULT ) PURE;
    STDMETHOD( PropertyFailure )( THIS_ IShellItem *, const PROPERTYKEY *, HRESULT ) PURE;
};
#endif

/* ITransferSource interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   ITransferSource
DECLARE_INTERFACE_( ITransferSource, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITransferSource methods */
    STDMETHOD( Advise )( THIS_ ITransferAdviseSink *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetProperties )( THIS_ IPropertyChangeArray * ) PURE;
    STDMETHOD( OpenItem )( THIS_ IShellItem *, TRANSFER_SOURCE_FLAGS, REFIID, void ** ) PURE;
    STDMETHOD( MoveItem )( THIS_ IShellItem *, IShellItem *, LPCWSTR, TRANSFER_SOURCE_FLAGS, IShellItem ** ) PURE;
    STDMETHOD( RecycleItem )( THIS_ IShellItem *, IShellItem *, TRANSFER_SOURCE_FLAGS, IShellItem ** ) PURE;
    STDMETHOD( RemoveItem )( THIS_ IShellItem *, TRANSFER_SOURCE_FLAGS ) PURE;
    STDMETHOD( RenameItem )( THIS_ IShellItem *, LPCWSTR, TRANSFER_SOURCE_FLAGS, IShellItem ** ) PURE;
    STDMETHOD( LinkItem )( THIS_ IShellItem *, IShellItem *, LPCWSTR, TRANSFER_SOURCE_FLAGS, IShellItem ** ) PURE;
    STDMETHOD( ApplyPropertiesToItem )( THIS_ IShellItem *, IShellItem ** ) PURE;
    STDMETHOD( GetDefaultDestinationName )( THIS_ IShellItem *, IShellItem *, LPWSTR * ) PURE;
    STDMETHOD( EnterFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( LeaveFolder )( THIS_ IShellItem * ) PURE;
};
#endif

/* IEnumResources interface */
#undef INTERFACE
#define INTERFACE   IEnumResources
DECLARE_INTERFACE_( IEnumResources, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumResources methods */
    STDMETHOD( Next )( THIS_ ULONG, SHELL_ITEM_RESOURCE *, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumResources ** ) PURE;
};

/* IShellItemResources interface */
#undef INTERFACE
#define INTERFACE   IShellItemResources
DECLARE_INTERFACE_( IShellItemResources, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellItemResources methods */
    STDMETHOD( GetAttributes )( THIS_ DWORD * ) PURE;
    STDMETHOD( GetSize )( THIS_ ULONGLONG * ) PURE;
    STDMETHOD( GetTimes )( THIS_ FILETIME *, FILETIME *, FILETIME * ) PURE;
    STDMETHOD( SetTimes )( THIS_ const FILETIME *, const FILETIME *, const FILETIME * ) PURE;
    STDMETHOD( GetResourceDescription )( THIS_ const SHELL_ITEM_RESOURCE *, LPWSTR * ) PURE;
    STDMETHOD( EnumResources )( THIS_ IEnumResources ** ) PURE;
    STDMETHOD( SupportsResource )( THIS_ const SHELL_ITEM_RESOURCE * ) PURE;
    STDMETHOD( OpenResource )( THIS_ const SHELL_ITEM_RESOURCE *, REFIID, void ** ) PURE;
    STDMETHOD( CreateResource )( THIS_ const SHELL_ITEM_RESOURCE *, REFIID, void ** ) PURE;
    STDMETHOD( MarkForDelete )( THIS ) PURE;
};

/* ITransferDestination interface */
#undef INTERFACE
#define INTERFACE   ITransferDestination
DECLARE_INTERFACE_( ITransferDestination, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITransferDestination methods */
    STDMETHOD( Advise )( THIS_ ITransferAdviseSink *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( CreateItem )( THIS_ LPCWSTR, DWORD, ULONGLONG, TRANSFER_SOURCE_FLAGS, REFIID, void **, REFIID, void ** ) PURE;
};

/* IStreamAsync interface */
#undef INTERFACE
#define INTERFACE   IStreamAsync
DECLARE_INTERFACE_( IStreamAsync, IStream ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* ISequentialStream methods */
    STDMETHOD( Read )( THIS_ void *, ULONG, ULONG * ) PURE;
    STDMETHOD( Write )( THIS_ void const *, ULONG, ULONG * ) PURE;
    
    /* IStream methods */
    STDMETHOD( Seek )( THIS_ LARGE_INTEGER, DWORD, ULARGE_INTEGER * ) PURE;
    STDMETHOD( SetSize )( THIS_ ULARGE_INTEGER ) PURE;
    STDMETHOD( CopyTo )( THIS_ IStream *, ULARGE_INTEGER, ULARGE_INTEGER *, ULARGE_INTEGER * ) PURE;
    STDMETHOD( Commit )( THIS_ DWORD ) PURE;
    STDMETHOD( Revert )( THIS ) PURE;
    STDMETHOD( LockRegion )( THIS_ ULARGE_INTEGER, ULARGE_INTEGER, DWORD ) PURE;
    STDMETHOD( UnlockRegion )( THIS_ ULARGE_INTEGER, ULARGE_INTEGER, DWORD ) PURE;
    STDMETHOD( Stat )( THIS_ STATSTG *, DWORD ) PURE;
    STDMETHOD( Clone )( THIS_ IStream ** ) PURE;

    /* IStreamAsync methods */
    STDMETHOD( ReadAsync )( THIS_ void *, DWORD, LPDWORD, LPOVERLAPPED ) PURE;
    STDMETHOD( WriteAsync )( THIS_ const void *, DWORD, LPDWORD, LPOVERLAPPED ) PURE;
    STDMETHOD( OverlappedResult )( THIS_ LPOVERLAPPED, LPDWORD, BOOL ) PURE;
    STDMETHOD( CancelIo )( THIS ) PURE;
};

/* IStreamUnbufferedInfo interface */
#undef INTERFACE
#define INTERFACE   IStreamUnbufferedInfo
DECLARE_INTERFACE_( IStreamUnbufferedInfo, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IStreamUnbufferedInfo methods */
    STDMETHOD( GetSectorSize )( THIS_ ULONG * ) PURE;
};

/* IFileOperationProgressSink interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IFileOperationProgressSink
DECLARE_INTERFACE_( IFileOperationProgressSink, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileOperationProgressSink methods */
    STDMETHOD( StartOperations )( THIS ) PURE;
    STDMETHOD( FinishOperations )( THIS_ HRESULT ) PURE;
    STDMETHOD( PreRenameItem )( THIS_ DWORD, IShellItem *, LPCWSTR ) PURE;
    STDMETHOD( PostRenameItem )( THIS_ DWORD, IShellItem *, LPCWSTR, HRESULT, IShellItem * ) PURE;
    STDMETHOD( PreMoveItem )( THIS_ DWORD, IShellItem *, IShellItem *, LPCWSTR ) PURE;
    STDMETHOD( PostMoveItem )( THIS_ DWORD, IShellItem *, IShellItem *, LPCWSTR, HRESULT, IShellItem * ) PURE;
    STDMETHOD( PreCopyItem )( THIS_ DWORD, IShellItem *, IShellItem *, LPCWSTR ) PURE;
    STDMETHOD( PostCopyItem )( THIS_ DWORD, IShellItem *, IShellItem *, LPCWSTR, HRESULT, IShellItem * ) PURE;
    STDMETHOD( PreDeleteItem )( THIS_ DWORD, IShellItem * ) PURE;
    STDMETHOD( PostDeleteItem )( THIS_ DWORD, IShellItem *, HRESULT, IShellItem * ) PURE;
    STDMETHOD( PreNewItem )( THIS_ DWORD, IShellItem *, LPCWSTR ) PURE;
    STDMETHOD( PostNewItem )( THIS_ DWORD, IShellItem *, LPCWSTR, LPCWSTR, DWORD, HRESULT, IShellItem * ) PURE;
    STDMETHOD( UpdateProgress )( THIS_ UINT, UINT ) PURE;
    STDMETHOD( ResetTimer )( THIS ) PURE;
    STDMETHOD( PauseTimer )( THIS ) PURE;
    STDMETHOD( ResumeTimer )( THIS ) PURE;
};
#endif

/* IShellItemArray interface */
#undef INTERFACE
#define INTERFACE   IShellItemArray
DECLARE_INTERFACE_( IShellItemArray, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellItemArray methods */
    STDMETHOD( BindToHandler )( THIS_ IBindCtx *, REFGUID, REFIID, void ** ) PURE;
    STDMETHOD( GetPropertyStore )( THIS_ GETPROPERTYSTOREFLAGS, REFIID, void ** ) PURE;
    STDMETHOD( GetPropertyDescriptionList )( THIS_ REFPROPERTYKEY, REFIID, void ** ) PURE;
    STDMETHOD( GetAttributes )( THIS_ SIATTRIBFLAGS, SFGAOF, SFGAOF * ) PURE;
    STDMETHOD( GetCount )( THIS_ DWORD * ) PURE;
    STDMETHOD( GetItemAt )( THIS_ DWORD, IShellItem ** ) PURE;
    STDMETHOD( EnumItems )( THIS_ IEnumShellItems ** ) PURE;
};

/* IInitializeWithItem interface */
#undef INTERFACE
#define INTERFACE   IInitializeWithItem
DECLARE_INTERFACE_( IInitializeWithItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInitializeWithItem methods */
    STDMETHOD( Initialize )( THIS_ IShellItem *, DWORD ) PURE;
};

/* IObjectWithSelection interface */
#undef INTERFACE
#define INTERFACE   IObjectWithSelection
DECLARE_INTERFACE_( IObjectWithSelection, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectWithSelection methods */
    STDMETHOD( SetSelection )( THIS_ IShellItemArray * ) PURE;
    STDMETHOD( GetSelection )( THIS_ REFIID, void ** ) PURE;
};

/* IObjectWithBackReferences interface */
#undef INTERFACE
#define INTERFACE   IObjectWithBackReferences
DECLARE_INTERFACE_( IObjectWithBackReferences, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectWithBackReferences methods */
    STDMETHOD( RemoveBackReferences )( THIS ) PURE;
};

/* IPropertyUI interface */
#undef INTERFACE
#define INTERFACE   IPropertyUI
DECLARE_INTERFACE_( IPropertyUI, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IPropertyUI methods */
    STDMETHOD( ParsePropertyName )( THIS_ LPCWSTR, FMTID *, PROPID *, ULONG * ) PURE;
    STDMETHOD( GetCanonicalName )( THIS_ REFFMTID, PROPID, LPWSTR, DWORD ) PURE;
    STDMETHOD( GetDisplayName )( THIS_ REFFMTID, PROPID, PROPERTYUI_NAME_FLAGS, LPWSTR, DWORD ) PURE;
    STDMETHOD( GetPropertyDescription )( THIS_ REFFMTID, PROPID, LPWSTR, DWORD ) PURE;
    STDMETHOD( GetDefaultWidth )( THIS_ REFFMTID, PROPID, ULONG * ) PURE;
    STDMETHOD( GetFlags )( THIS_ REFFMTID, PROPID, PROPERTYUI_FLAGS * ) PURE;
    STDMETHOD( FormatForDisplay )( THIS_ REFFMTID, PROPID, const PROPVARIANT *, PROPERTYUI_FORMAT_FLAGS, LPWSTR, DWORD ) PURE;
    STDMETHOD( GetHelpInfo )( THIS_ REFFMTID, PROPID, LPWSTR, DWORD, UINT * ) PURE;
};

/* ICategoryProvider interface */
#if (_WIN32_IE >= 0x0500)
#undef INTERFACE
#define INTERFACE   ICategoryProvider
DECLARE_INTERFACE_( ICategoryProvider, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICategoryProvider methods */
    STDMETHOD( CanCategorizeOnSCID )( THIS_ const SHCOLUMNID * ) PURE;
    STDMETHOD( GetDefaultCategory )( THIS_ GUID *, SHCOLUMNID * ) PURE;
    STDMETHOD( GetCategoryForSCID )( THIS_ const SHCOLUMNID *, GUID * ) PURE;
    STDMETHOD( EnumCategories )( THIS_ IEnumGUID ** ) PURE;
    STDMETHOD( GetCategoryName )( THIS_ const GUID *, LPWSTR, UINT ) PURE;
    STDMETHOD( CreateCategory )( THIS_ const GUID *, REFIID, void ** ) PURE;
};
#endif

/* ICategorizer interface */
#if (_WIN32_IE >= 0x0500)
#undef INTERFACE
#define INTERFACE   ICategorizer
DECLARE_INTERFACE_( ICategorizer, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICategorizer methods */
    STDMETHOD( GetDescription )( THIS_ LPWSTR, UINT ) PURE;
    STDMETHOD( GetCategory )( THIS_ UINT, PCUITEMID_CHILD_ARRAY, DWORD * ) PURE;
    STDMETHOD( GetCategoryInfo )( THIS_ DWORD, CATEGORY_INFO * ) PURE;
    STDMETHOD( CompareCategory )( THIS_ CATSORT_FLAGS, DWORD, DWORD ) PURE;
};
#endif

/* IDropTargetHelper interface */
#if (NTDDI_VERSION >= 0x05000000)
#undef INTERFACE
#define INTERFACE   IDropTargetHelper
DECLARE_INTERFACE_( IDropTargetHelper, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDropTargetHelper methods */
    STDMETHOD( DragEnter )( THIS_ HWND, IDataObject *, POINT *, DWORD ) PURE;
    STDMETHOD( DragLeave )( THIS ) PURE;
    STDMETHOD( DragOver )( THIS_ POINT *, DWORD ) PURE;
    STDMETHOD( Drop )( THIS_ IDataObject *, POINT *, DWORD ) PURE;
    STDMETHOD( Show )( THIS_ BOOL ) PURE;
};
#endif

/* IDragSourceHelper interface */
#if (NTDDI_VERSION >= 0x05000000)
#undef INTERFACE
#define INTERFACE   IDragSourceHelper
DECLARE_INTERFACE_( IDragSourceHelper, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDragSourceHelper methods */
    STDMETHOD( InitializeFromBitmap )( THIS_ LPSHDRAGIMAGE, IDataObject * ) PURE;
    STDMETHOD( InitializeFromWindow )( THIS_ HWND, POINT *, IDataObject * ) PURE;
};
#endif

/* IDragSourceHelper2 interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IDragSourceHelper2
DECLARE_INTERFACE_( IDragSourceHelper2, IDragSourceHelper ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDragSourceHelper methods */
    STDMETHOD( InitializeFromBitmap )( THIS_ LPSHDRAGIMAGE, IDataObject * ) PURE;
    STDMETHOD( InitializeFromWindow )( THIS_ HWND, POINT *, IDataObject * ) PURE;

    /* IDragSourceHelper2 methods */
    STDMETHOD( SetFlags )( THIS_ DWORD ) PURE;
};
#endif

/* IShellLink interface */
#undef INTERFACE
#define INTERFACE   IShellLinkA
DECLARE_INTERFACE_( IShellLinkA, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellLinkA methods */
    STDMETHOD( GetPath )( THIS_ LPSTR, int, WIN32_FIND_DATAA *, DWORD ) PURE;
    STDMETHOD( GetIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( SetIDList )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
    STDMETHOD( GetDescription )( THIS_ LPSTR, int ) PURE;
    STDMETHOD( SetDescription )( THIS_ LPCSTR ) PURE;
    STDMETHOD( GetWorkingDirectory )( THIS_ LPSTR, int ) PURE;
    STDMETHOD( SetWorkingDirectory )( THIS_ LPCSTR ) PURE;
    STDMETHOD( GetArguments )( THIS_ LPSTR, int ) PURE;
    STDMETHOD( SetArguments )( THIS_ LPCSTR ) PURE;
    STDMETHOD( GetHotKey )( THIS_ WORD * ) PURE;
    STDMETHOD( SetHotKey )( THIS_ WORD ) PURE;
    STDMETHOD( GetShowCmd )( THIS_ int * ) PURE;
    STDMETHOD( SetShowCmd )( THIS_ int ) PURE;
    STDMETHOD( GetIconLocation )( THIS_ LPSTR, int, int * ) PURE;
    STDMETHOD( SetIconLocation )( THIS_ LPCSTR, int ) PURE;
    STDMETHOD( SetRelativePath )( THIS_ LPCSTR, DWORD ) PURE;
    STDMETHOD( Resolve )( THIS_ HWND, DWORD ) PURE;
    STDMETHOD( SetPath )( THIS_ LPCSTR ) PURE;
};
#undef INTERFACE
#define INTERFACE   IShellLinkW
DECLARE_INTERFACE_( IShellLinkW, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellLinkW methods */
    STDMETHOD( GetPath )( THIS_ LPWSTR, int, WIN32_FIND_DATAW *, DWORD ) PURE;
    STDMETHOD( GetIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( SetIDList )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
    STDMETHOD( GetDescription )( THIS_ LPWSTR, int ) PURE;
    STDMETHOD( SetDescription )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetWorkingDirectory )( THIS_ LPWSTR, int ) PURE;
    STDMETHOD( SetWorkingDirectory )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetArguments )( THIS_ LPWSTR, int ) PURE;
    STDMETHOD( SetArguments )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetHotKey )( THIS_ WORD * ) PURE;
    STDMETHOD( GetHotKey )( THIS_ WORD ) PURE;
    STDMETHOD( GetShowCmd )( THIS_ int * ) PURE;
    STDMETHOD( SetShowCmd )( THIS_ int ) PURE;
    STDMETHOD( GetIconLocation )( THIS_ LPWSTR, int, int * ) PURE;
    STDMETHOD( SetIconLocation )( THIS_ LPCWSTR, int ) PURE;
    STDMETHOD( SetRelativePath )( THIS_ LPCWSTR, DWORD ) PURE;
    STDMETHOD( Resolve )( THIS_ HWND, DWORD ) PURE;
    STDMETHOD( SetPath )( THIS_ LPCWSTR ) PURE;
};
#ifdef UNICODE
    #define IShellLink  IShellLinkW
#else
    #define IShellLink  IShellLinkA
#endif

/* IShellLinkDataList interface */
#undef INTERFACE
#define INTERFACE   IShellLinkDataList
DECLARE_INTERFACE_( IShellLinkDataList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellLinkDataList methods */
    STDMETHOD( AddDataBlock )( THIS_ void * ) PURE;
    STDMETHOD( CopyDataBlock )( THIS_ DWORD, void ** ) PURE;
    STDMETHOD( RemoveDataBlock )( THIS_ DWORD ) PURE;
    STDMETHOD( GetFlags )( THIS_ DWORD * ) PURE;
    STDMETHOD( SetFlags )( THIS_ DWORD ) PURE;
};

/* IResolveShellLink interface */
#if (NTDDI_VERSION >= 0x05000000)
#undef INTERFACE
#define INTERFACE   IResolveShellLink
DECLARE_INTERFACE_( IResolveShellLink, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IResolveShellLink methods */
    STDMETHOD( ResolveShellLink )( THIS_ IUnknown *, HWND, DWORD ) PURE;
};
#endif

/* IActionProgressDialog interface */
#undef INTERFACE
#define INTERFACE   IActionProgressDialog
DECLARE_INTERFACE_( IActionProgressDialog, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IActionProgressDialog methods */
    STDMETHOD( Initialize )( THIS_ SPINITF, LPCWSTR, LPCWSTR ) PURE;
    STDMETHOD( Stop )( THIS ) PURE;
};

/* IHWEventHandler interface */
#undef INTERFACE
#define INTERFACE   IHWEventHandler
DECLARE_INTERFACE_( IHWEventHandler, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IHWEventHandler methods */
    STDMETHOD( Initialize )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( HandleEvent )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR ) PURE;
    STDMETHOD( HandleEventWithContent )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR, IDataObject * ) PURE;
};

/* IHWEventHandler2 interface */
#undef INTERFACE
#define INTERFACE   IHWEventHandler2
DECLARE_INTERFACE_( IHWEventHandler2, IHWEventHandler ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IHWEventHandler methods */
    STDMETHOD( Initialize )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( HandleEvent )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR ) PURE;
    STDMETHOD( HandleEventWithContent )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR, IDataObject * ) PURE;

    /* IHWEventHandler2 methods */
    STDMETHOD( HandleEventWithHWND )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR, HWND ) PURE;
};

/* IQueryCancelAutoPlay interface */
#undef INTERFACE
#define INTERFACE   IQueryCancelAutoPlay
DECLARE_INTERFACE_( IQueryCancelAutoPlay, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IQueryCancelAutoPlay methods */
    STDMETHOD( AllowAutoPlay )( THIS_ LPCWSTR, DWORD, LPCWSTR, DWORD ) PURE;
};

/* IDynamicHWHandler interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IDynamicHWHandler
DECLARE_INTERFACE_( IDynamicHWHandler, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDynamicHWHandler methods */
    STDMETHOD( GetDynamicInfo )( THIS_ LPCWSTR, DWORD, LPWSTR * ) PURE;
};
#endif

/* IActionProgress interface */
#undef INTERFACE
#define INTERFACE   IActionProgress
DECLARE_INTERFACE_( IActionProgress, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IActionProgress methods */
    STDMETHOD( Begin )( THIS_ SPACTION, SPBEGINF ) PURE;
    STDMETHOD( UpdateProgress )( THIS_ ULONGLONG, ULONGLONG ) PURE;
    STDMETHOD( UpdateText )( THIS_ SPTEXT, LPCWSTR, BOOL ) PURE;
    STDMETHOD( QueryCancel )( THIS_ BOOL * ) PURE;
    STDMETHOD( ResetCancel )( THIS ) PURE;
    STDMETHOD( End )( THIS ) PURE;
};

/* IShellExtInit interface */
#undef INTERFACE
#define INTERFACE   IShellExtInit
DECLARE_INTERFACE_( IShellExtInit, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellExtInit methods */
    STDMETHOD( Initialize )( THIS_ PCIDLIST_ABSOLUTE, IDataObject *, HKEY ) PURE;
};
typedef IShellExtInit   *LPSHELLEXTINIT;

/* IShellPropSheetExt interface */
#undef INTERFACE
#define INTERFACE   IShellPropSheetExt
DECLARE_INTERFACE_( IShellPropSheetExt, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellPropSheetExt methods */
    STDMETHOD( AddPages )( THIS_ LPFNSVADDPROPSHEETPAGE, LPARAM ) PURE;
    STDMETHOD( ReplacePage )( THIS_ EXPPS, LPFNSVADDPROPSHEETPAGE, LPARAM ) PURE;
};
typedef IShellPropSheetExt  *LPSHELLPROPSHEETEXT;

/* IRemoteComputer interface */
#undef INTERFACE
#define INTERFACE   IRemoteComputer
DECLARE_INTERFACE_( IRemoteComputer, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRemoteComputer methods */
    STDMETHOD( Initialize )( THIS_ LPCWSTR, BOOL ) PURE;
};

/* IQueryContinue interface */
#undef INTERFACE
#define INTERFACE   IQueryContinue
DECLARE_INTERFACE_( IQueryContinue, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IQueryContinue methods */
    STDMETHOD( QueryContinue )( THIS ) PURE;
};

/* IObjectWithCancelEvent interface */
#undef INTERFACE
#define INTERFACE   IObjectWithCancelEvent
DECLARE_INTERFACE_( IObjectWithCancelEvent, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectWithCancelEvent methods */
    STDMETHOD( GetCancelEvent )( THIS_ HANDLE * ) PURE;
};

/* IUserNotification interface */
#undef INTERFACE
#define INTERFACE   IUserNotification
DECLARE_INTERFACE_( IUserNotification, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IUserNotification methods */
    STDMETHOD( SetBalloonInfo )( THIS_ LPCWSTR, LPCWSTR, DWORD ) PURE;
    STDMETHOD( SetBalloonRetry )( THIS_ DWORD, DWORD, UINT ) PURE;
    STDMETHOD( SetIconInfo )( THIS_ HICON, LPCWSTR ) PURE;
    STDMETHOD( Show )( THIS_ IQueryContinue *, DWORD ) PURE;
    STDMETHOD( PlaySound )( THIS_ LPCWSTR ) PURE;
};

/* IUserNotificationCallback interface */
#undef INTERFACE
#define INTERFACE   IUserNotificationCallback
DECLARE_INTERFACE_( IUserNotificationCallback, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IUserNotificationCallback methods */
    STDMETHOD( OnBalloonUserClick )( THIS_ POINT * ) PURE;
    STDMETHOD( OnLeftClick )( THIS_ POINT * ) PURE;
    STDMETHOD( OnContextMenu )( THIS_ POINT * ) PURE;
};

/* IUserNotification2 interface */
#undef INTERFACE
#define INTERFACE   IUserNotification2
DECLARE_INTERFACE_( IUserNotification2, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IUserNotification methods */
    STDMETHOD( SetBalloonInfo )( THIS_ LPCWSTR, LPCWSTR, DWORD ) PURE;
    STDMETHOD( SetBalloonRetry )( THIS_ DWORD, DWORD, UINT ) PURE;
    STDMETHOD( SetIconInfo )( THIS_ HICON, LPCWSTR ) PURE;
    STDMETHOD( Show )( THIS_ IQueryContinue *, DWORD, IUserNotificationCallback * ) PURE;
    STDMETHOD( PlaySound )( THIS_ LPCWSTR ) PURE;
};

/* IItemNameLimits interface */
#undef INTERFACE
#define INTERFACE   IItemNameLimits
DECLARE_INTERFACE_( IItemNameLimits, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IItemNameLimits methods */
    STDMETHOD( GetValidCharacters )( THIS_ LPWSTR *, LPWSTR * ) PURE;
    STDMETHOD( GetMaxLength )( THIS_ LPCWSTR, int * ) PURE;
};

/* ISearchFolderItemFactory interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   ISearchFolderItemFactory
DECLARE_INTERFACE_( ISearchFolderItemFactory, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ISearchFolderItemFactory methods */
    STDMETHOD( SetDisplayName )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFolderTypeID )( THIS_ FOLDERTYPEID ) PURE;
    STDMETHOD( SetFolderLogicalViewMode )( THIS_ FOLDERLOGICALVIEWMODE ) PURE;
    STDMETHOD( SetIconSize )( THIS_ int ) PURE;
    STDMETHOD( SetVisibleColumns )( THIS_ UINT, PROPERTYKEY * ) PURE;
    STDMETHOD( SetSortColumns )( THIS_ UINT, SORTCOLUMN * ) PURE;
    STDMETHOD( SetGroupColumn )( THIS_ REFPROPERTYKEY ) PURE;
    STDMETHOD( SetStacks )( THIS_ UINT, PROPERTYKEY * ) PURE;
    STDMETHOD( SetScope )( THIS_ IShellItemArray * ) PURE;
    STDMETHOD( SetCondition )( THIS_ ICondition * ) PURE;
    STDMETHOD( GetShellItem )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( GetIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
};
#endif

/* IExtractImage interface */
#if (_WIN32_IE >= 0x0400)
#undef INTERFACE
#define INTERFACE   IExtractImage
DECLARE_INTERFACE_( IExtractImage, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExtractImage methods */
    STDMETHOD( GetLocation )( THIS_ LPWSTR, DWORD, DWORD *, const SIZE *, DWORD, DWORD * ) PURE;
    STDMETHOD( Extract )( THIS_ HBITMAP * ) PURE;
};
typedef IExtractImage   *LPEXTRACTIMAGE;
#endif

/* IExtractImage2 interface */
#if (_WIN32_IE >= 0x0500)
#undef INTERFACE
#define INTERFACE   IExtractImage2
DECLARE_INTERFACE_( IExtractImage2, IExtractImage ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExtractImage methods */
    STDMETHOD( GetLocation )( THIS_ LPWSTR, DWORD, DWORD *, const SIZE *, DWORD, DWORD * ) PURE;
    STDMETHOD( Extract )( THIS_ HBITMAP * ) PURE;

    /* IExtractImage2 methods */
    STDMETHOD( GetDateStamp )( THIS_ FILETIME * ) PURE;
};
typedef IExtractImage2  *LPEXTRACTIMAGE2;
#endif

/* IThumbnailHandlerFactory interface */
#if (_WIN32_IE >= 0x0500)
#undef INTERFACE
#define INTERFACE   IThumbnailHandlerFactory
DECLARE_INTERFACE_( IThumbnailHandlerFactory, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IThumbnailHandlerFactory methods */
    STDMETHOD( GetThumbnailHandler )( THIS_ PCUITEMID_CHILD, IBindCtx *, REFIID, void ** ) PURE;
};
#endif

/* IParentAndItem interface */
#if (_WIN32_IE >= 0x0500)
#undef INTERFACE
#define INTERFACE   IParentAndItem
DECLARE_INTERFACE_( IParentAndItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IParentAndItem methods */
    STDMETHOD( SetParentAndItem )( THIS_ PCIDLIST_ABSOLUTE, IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( GetParentAndItem )( THIS_ PIDLIST_ABSOLUTE *, IShellFolder **, PITEMID_CHILD * ) PURE;
};
#endif

/* IDockingWindow interface */
#undef INTERFACE
#define INTERFACE   IDockingWindow
DECLARE_INTERFACE_( IDockingWindow, IOleWindow ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IDockingWindow methods */
    STDMETHOD( ShowDW )( THIS_ BOOL ) PURE;
    STDMETHOD( CloseDW )( THIS_ DWORD ) PURE;
    STDMETHOD( ResizeBorderDW )( THIS_ LPCRECT, IUnknown *, BOOL ) PURE;
};

/* IDeskBand interface */
#undef INTERFACE
#define INTERFACE   IDeskBand
DECLARE_INTERFACE_( IDeskBand, IDockingWindow ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IDockingWindow methods */
    STDMETHOD( ShowDW )( THIS_ BOOL ) PURE;
    STDMETHOD( CloseDW )( THIS_ DWORD ) PURE;
    STDMETHOD( ResizeBorderDW )( THIS_ LPCRECT, IUnknown *, BOOL ) PURE;

    /* IDeskBand methods */
    STDMETHOD( GetBandInfo )( THIS_ DWORD, DWORD, DESKBANDINFO * ) PURE;
};

/* IDeskBandInfo interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IDeskBandInfo
DECLARE_INTERFACE_( IDeskBandInfo, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDeskBandInfo methods */
    STDMETHOD( GetDefaultBandWidth )( THIS_ DWORD, DWORD, int * ) PURE;
};
#endif

/* IDeskBand2 interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IDeskBand2
DECLARE_INTERFACE_( IDeskBand2, IDeskBand ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IDockingWindow methods */
    STDMETHOD( ShowDW )( THIS_ BOOL ) PURE;
    STDMETHOD( CloseDW )( THIS_ DWORD ) PURE;
    STDMETHOD( ResizeBorderDW )( THIS_ LPCRECT, IUnknown *, BOOL ) PURE;

    /* IDeskBand methods */
    STDMETHOD( GetBandInfo )( THIS_ DWORD, DWORD, DESKBANDINFO * ) PURE;

    /* DeskBand2 methods */
    STDMETHOD( CanRenderComposited )( THIS_ BOOL * ) PURE;
    STDMETHOD( SetCompositionState )( THIS_ BOOL ) PURE;
    STDMETHOD( GetCompositionState )( THIS_ BOOL * ) PURE;
};
#endif

/* ITaskbarList interface */
#undef INTERFACE
#define INTERFACE   ITaskbarList
DECLARE_INTERFACE_( ITaskbarList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITaskbarList methods */
    STDMETHOD( HrInit )( THIS ) PURE;
    STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;
};

/* ITaskbarList2 interface */
#undef INTERFACE
#define INTERFACE   ITaskbarList2
DECLARE_INTERFACE_( ITaskbarList2, ITaskbarList ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITaskbarList methods */
    STDMETHOD( HrInit )( THIS ) PURE;
    STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;

    /* ITaskbarList2 methods */
    STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;
};

/* ITaskbarList3 interface */
#undef INTERFACE
#define INTERFACE   ITaskbarList3
DECLARE_INTERFACE_( ITaskbarList3, ITaskbarList2 ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITaskbarList methods */
    STDMETHOD( HrInit )( THIS ) PURE;
    STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;

    /* ITaskbarList2 methods */
    STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;

    /* ITaskbarList3 methods */
    STDMETHOD( SetProgressValue )( THIS_ HWND, ULONGLONG, ULONGLONG ) PURE;
    STDMETHOD( SetProgressState )( THIS_ HWND, TBPFLAG ) PURE;
    STDMETHOD( RegisterTab )( THIS_ HWND, HWND ) PURE;
    STDMETHOD( UnregisterTab )( THIS_ HWND ) PURE;
    STDMETHOD( SetTabOrder )( THIS_ HWND, HWND ) PURE;
    STDMETHOD( SetTabActive )( THIS_ HWND, HWND, DWORD ) PURE;
    STDMETHOD( ThumbBarAddButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    STDMETHOD( ThumbBarUpdateButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    STDMETHOD( ThumbBarSetImageList )( THIS_ HWND, HIMAGELIST ) PURE;
    STDMETHOD( SetOverlayIcon )( THIS_ HWND, HICON, LPCWSTR ) PURE;
    STDMETHOD( SetThumbnailTooltip )( THIS_ HWND, LPCWSTR ) PURE;
    STDMETHOD( SetThumbnailClip )( THIS_ HWND, RECT * ) PURE;
};

/* ITaskbarList4 interface */
#undef INTERFACE
#define INTERFACE   ITaskbarList4
DECLARE_INTERFACE_( ITaskbarList4, ITaskbarList3 ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITaskbarList methods */
    STDMETHOD( HrInit )( THIS ) PURE;
    STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;

    /* ITaskbarList2 methods */
    STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;

    /* ITaskbarList3 methods */
    STDMETHOD( SetProgressValue )( THIS_ HWND, ULONGLONG, ULONGLONG ) PURE;
    STDMETHOD( SetProgressState )( THIS_ HWND, TBPFLAG ) PURE;
    STDMETHOD( RegisterTab )( THIS_ HWND, HWND ) PURE;
    STDMETHOD( UnregisterTab )( THIS_ HWND ) PURE;
    STDMETHOD( SetTabOrder )( THIS_ HWND, HWND ) PURE;
    STDMETHOD( SetTabActive )( THIS_ HWND, HWND, DWORD ) PURE;
    STDMETHOD( ThumbBarAddButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    STDMETHOD( ThumbBarUpdateButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    STDMETHOD( ThumbBarSetImageList )( THIS_ HWND, HIMAGELIST ) PURE;
    STDMETHOD( SetOverlayIcon )( THIS_ HWND, HICON, LPCWSTR ) PURE;
    STDMETHOD( SetThumbnailTooltip )( THIS_ HWND, LPCWSTR ) PURE;
    STDMETHOD( SetThumbnailClip )( THIS_ HWND, RECT * ) PURE;

    /* ITaskbarList4 methods */
    STDMETHOD( SetTabProperties )( THIS_ HWND, STPFLAG ) PURE;
};

/* IStartMenuPinnedList interface */
#undef INTERFACE
#define INTERFACE   IStartMenuPinnedList
DECLARE_INTERFACE_( IStartMenuPinnedList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IStartMenuPinnedList methods */
    STDMETHOD( RemoveFromList )( THIS_ IShellItem * ) PURE;
};

/* ICDBurn interface */
#undef INTERFACE
#define INTERFACE   ICDBurn
DECLARE_INTERFACE_( ICDBurn, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICDBurn methods */
    STDMETHOD( GetRecorderDriverLetter )( THIS_ LPWSTR, UINT ) PURE;
    STDMETHOD( Burn )( THIS_ HWND ) PURE;
    STDMETHOD( HasRecordableDrive )( THIS_ BOOL * ) PURE;
};

/* IWizardSite interface */
#undef INTERFACE
#define INTERFACE   IWizardSite
DECLARE_INTERFACE_( IWizardSite, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IWizardSite methods */
    STDMETHOD( GetPreviousPage )( THIS_ HPROPSHEETPAGE * ) PURE;
    STDMETHOD( GetNextPage )( THIS_ HPROPSHEETPAGE * ) PURE;
    STDMETHOD( GetCancelledPage )( THIS_ HPROPSHEETPAGE * ) PURE;
};

/* IWizardExtension interface */
#undef INTERFACE
#define INTERFACE   IWizardExtension
DECLARE_INTERFACE_( IWizardExtension, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IWizardExtension methods */
    STDMETHOD( AddPages )( THIS_ HPROPSHEETPAGE *, UINT, UINT * ) PURE;
    STDMETHOD( GetFirstPage )( THIS_ HPROPSHEETPAGE * ) PURE;
    STDMETHOD( GetLastPage )( THIS_ HPROPSHEETPAGE * ) PURE;
};

/* IWebWizardExtension interface */
#undef INTERFACE
#define INTERFACE   IWebWizardExtension
DECLARE_INTERFACE_( IWebWizardExtension, IWizardExtension ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IWizardExtension methods */
    STDMETHOD( AddPages )( THIS_ HPROPSHEETPAGE *, UINT, UINT * ) PURE;
    STDMETHOD( GetFirstPage )( THIS_ HPROPSHEETPAGE * ) PURE;
    STDMETHOD( GetLastPage )( THIS_ HPROPSHEETPAGE * ) PURE;

    /* IWebWizardExtension methods */
    STDMETHOD( SetInitialURL )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetErrorURL )( THIS_ LPCWSTR ) PURE;
};

/* IPublishingWizard interface */
#undef INTERFACE
#define INTERFACE   IPublishingWizard
DECLARE_INTERFACE_( IPublishingWizard, IWizardExtension ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IWizardExtension methods */
    STDMETHOD( AddPages )( THIS_ HPROPSHEETPAGE *, UINT, UINT * ) PURE;
    STDMETHOD( GetFirstPage )( THIS_ HPROPSHEETPAGE * ) PURE;
    STDMETHOD( GetLastPage )( THIS_ HPROPSHEETPAGE * ) PURE;

    /* IPublishingWizard methods */
    STDMETHOD( Initialize )( THIS_ IDataObject *, DWORD, LPCWSTR ) PURE;
    STDMETHOD( GetTransferManifest )( THIS_ HRESULT *, IXMLDOMDocument ** ) PURE;
};

/* IFolderViewHost interface */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IFolderViewHost
DECLARE_INTERFACE_( IFolderViewHost, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderViewHost methods */
    STDMETHOD( Initialize )( THIS_ HWND, IDataObject *, RECT * ) PURE;
};
#endif

/* IExplorerBrowserEvents interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IExplorerBrowserEvents
DECLARE_INTERFACE_( IExplorerBrowserEvents, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExplorerBrowserEvents methods */
    STDMETHOD( OnNavigationPending )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
    STDMETHOD( OnViewCreated )( THIS_ IShellView * ) PURE;
    STDMETHOD( OnNavigationCompleted )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
    STDMETHOD( OnNavigationFailed )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
};
#endif

/* IExplorerBrowser interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IExplorerBrowser
DECLARE_INTERFACE_( IExplorerBrowser, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExplorerBrowser methods */
    STDMETHOD( Initialize )( THIS_ HWND, const RECT *, const FOLDERSETTINGS * ) PURE;
    STDMETHOD( Destroy )( THIS ) PURE;
    STDMETHOD( SetRect )( THIS_ HDWP *, RECT ) PURE;
    STDMETHOD( SetPropertyBag )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetEmptyText )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFolderSettings )( THIS_ const FOLDERSETTINGS * ) PURE;
    STDMETHOD( Advise )( THIS_ IExplorerBrowserEvents *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetOptions )( THIS_ EXPLORER_BROWSER_OPTIONS ) PURE;
    STDMETHOD( GetOptions )( THIS_ EXPLORER_BROWSER_OPTIONS * ) PURE;
    STDMETHOD( BrowseToIDList )( THIS_ PCUIDLIST_RELATIVE, UINT ) PURE;
    STDMETHOD( BrowseToObject )( THIS_ IUnknown *, UINT ) PURE;
    STDMETHOD( FillFromObject )( THIS_ IUnknown *, EXPLORER_BROWSER_FILL_FLAGS ) PURE;
    STDMETHOD( RemoveAll )( THIS ) PURE;
    STDMETHOD( GetCurrentView )( THIS_ REFIID, void ** ) PURE;
};
#endif

/* IAccessibleObject interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IAccessibleObject
DECLARE_INTERFACE_( IAccessibleObject, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IAccessibleObject methods */
    STDMETHOD( SetAccessibleName )( THIS_ LPCWSTR ) PURE;
};
#endif

/* IResultsFolder interface */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IResultsFolder
DECLARE_INTERFACE_( IResultsFolder, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IResultsFolder methods */
    STDMETHOD( AddItem )( THIS_ IShellItem * ) PURE;
    STDMETHOD( AddIDList )( THIS_ PCIDLIST_ABSOLUTE, PITEMID_CHILD * ) PURE;
    STDMETHOD( RemoveItem )( THIS_ IShellItem * ) PURE;
    STDMETHOD( RemoveIDList )( THIS_ PCIDLIST_ABSOLUTE ) PURE;
    STDMETHOD( RemoveAll )( THIS ) PURE;
};
#endif

/* IEnumObjects interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IEnumObjects
DECLARE_INTERFACE_( IEnumObjects, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumObjects methods */
    STDMETHOD( Next )( THIS_ ULONG, REFIID, void **, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumObjects ** ) PURE;
};
#endif

/* IOperationsProgressDialog interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IOperationsProgressDialog
DECLARE_INTERFACE_( IOperationsProgressDialog, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IOperationsProgressDialog methods */
    STDMETHOD( StartProgressDialog )( THIS_ HWND, OPPROGDLGF ) PURE;
    STDMETHOD( StopProgressDialog )( THIS ) PURE;
    STDMETHOD( SetOperation )( THIS_ SPACTION ) PURE;
    STDMETHOD( SetMode )( THIS_ PDMODE ) PURE;
    STDMETHOD( UpdateProgress )( THIS_ ULONGLONG, ULONGLONG, ULONGLONG, ULONGLONG, ULONGLONG, ULONGLONG ) PURE;
    STDMETHOD( UpdateLocations )( THIS_ IShellItem *, IShellItem *, IShellItem * ) PURE;
    STDMETHOD( ResetTimer )( THIS ) PURE;
    STDMETHOD( PauseTimer )( THIS ) PURE;
    STDMETHOD( ResumeTimer )( THIS ) PURE;
    STDMETHOD( GetMilliseconds )( THIS_ ULONGLONG *, ULONGLONG * ) PURE;
    STDMETHOD( GetOperationStatus )( THIS_ PDOPSTATUS * ) PURE;
};
#endif

/* IIOCancelInformation interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IIOCancelInformation
DECLARE_INTERFACE_( IIOCancelInformation, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IIOCancelInformation methods */
    STDMETHOD( SetCancelInformation )( THIS_ DWORD, UINT ) PURE;
    STDMETHOD( GetCancelInformation )( THIS_ DWORD *, UINT * ) PURE;
};
#endif

/* IFileOperation interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IFileOperation
DECLARE_INTERFACE_( IFileOperation, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileOperation methods */
    STDMETHOD( Advise )( THIS_ IFileOperationProgressSink *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetOperationFlags )( THIS_ DWORD ) PURE;
    STDMETHOD( SetProgressMessage )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetProgressDialog )( THIS_ IOperationsProgressDialog * ) PURE;
    STDMETHOD( SetProperties )( THIS_ IPropertyChangeArray * ) PURE;
    STDMETHOD( SetOwnerWindow )( THIS_ HWND ) PURE;
    STDMETHOD( ApplyPropertiesToItem )( THIS_ IShellItem * ) PURE;
    STDMETHOD( ApplyPropertiesToItems )( THIS_ IUnknown * ) PURE;
    STDMETHOD( RenameItem )( THIS_ IShellItem *, LPCWSTR, IFileOperationProgressSink * ) PURE;
    STDMETHOD( RenameItems )( THIS_ IUnknown *, LPCWSTR ) PURE;
    STDMETHOD( MoveItem )( THIS_ IShellItem *, IShellItem *, LPCWSTR, IFileOperationProgressSink * ) PURE;
    STDMETHOD( MoveItems )( THIS_ IUnknown *, IShellItem * ) PURE;
    STDMETHOD( CopyItem )( THIS_ IShellItem *, IShellItem *, LPCWSTR, IFileOperationProgressSink * ) PURE;
    STDMETHOD( CopyItems )( THIS_ IUnknown *, IShellItem * ) PURE;
    STDMETHOD( DeleteItem )( THIS_ IShellItem *, IFileOperationProgressSink * ) PURE;
    STDMETHOD( DeleteItems )( THIS_ IUnknown * ) PURE;
    STDMETHOD( NewItem )( THIS_ IShellItem *, DWORD, LPCWSTR, LPCWSTR, IFileOperationProgressSink * ) PURE;
    STDMETHOD( PerformOperations )( THIS ) PURE;
    STDMETHOD( GetAnyOperationsAborted )( THIS_ BOOL * ) PURE;
};
#endif

/* IObjectProvider interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IObjectProvider
DECLARE_INTERFACE_( IObjectProvider, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectProvider methods */
    STDMETHOD( QueryObject )( THIS_ REFGUID, REFIID, void ** ) PURE;
};
#endif

/* INamespaceWalkCB interface */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   INamespaceWalkCB
DECLARE_INTERFACE_( INamespaceWalkCB, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* INamespaceWalkCB methods */
    STDMETHOD( FoundItem )( THIS_ IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( EnterFolder )( THIS_ IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( LeaveFolder )( THIS_ IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( InitializeProgressDialog )( THIS_ LPWSTR *, LPWSTR * ) PURE;
};
#endif

/* INamespaceWalkCB2 interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   INamespaceWalkCB2
DECLARE_INTERFACE_( INamespaceWalkCB2, INamespaceWalkCB ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* INamespaceWalkCB methods */
    STDMETHOD( FoundItem )( THIS_ IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( EnterFolder )( THIS_ IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( LeaveFolder )( THIS_ IShellFolder *, PCUITEMID_CHILD ) PURE;
    STDMETHOD( InitializeProgressDialog )( THIS_ LPWSTR *, LPWSTR * ) PURE;

    /* INamespaceWalkCB2 methods */
    STDMETHOD( WalkComplete )( THIS_ HRESULT ) PURE;
};
#endif

/* INamespaceWalk interface */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   INamespaceWalk
DECLARE_INTERFACE_( INamespaceWalk, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INamespaceWalk methods */
    STDMETHOD( Walk )( THIS_ IUnknown *, DWORD, int, INamespaceWalkCB * ) PURE;
    STDMETHOD( GetIDArrayResult )( THIS_ UINT *, PIDLIST_ABSOLUTE ** ) PURE;
};
#endif

/* IAutoCompleteDropDown interface */
#undef INTERFACE
#define INTERFACE   IAutoCompleteDropDown
DECLARE_INTERFACE_( IAutoCompleteDropDown, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IAutoCompleteDropDown methods */
    STDMETHOD( GetDropDownStatus )( THIS_ DWORD *, LPWSTR * ) PURE;
    STDMETHOD( ResetEnumerator )( THIS ) PURE;
};

/* IBandSite interface */
#if (_WIN32_IE >= 0x0400)
#undef INTERFACE
#define INTERFACE   IBandSite
DECLARE_INTERFACE_( IBandSite, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IBandSite methods */
    STDMETHOD( AddBand )( THIS_ IUnknown * ) PURE;
    STDMETHOD( EnumBands )( THIS_ UINT, DWORD * ) PURE;
    STDMETHOD( QueryBand )( THIS_ DWORD, IDeskBand **, DWORD *, LPWSTR, int ) PURE;
    STDMETHOD( SetBandState )( THIS_ DWORD, DWORD, DWORD ) PURE;
    STDMETHOD( RemoveBand )( THIS_ DWORD ) PURE;
    STDMETHOD( GetBandObject )( THIS_ DWORD, REFIID, void ** ) PURE;
    STDMETHOD( SetBandSiteInfo )( THIS_ const BANDSITEINFO * ) PURE;
    STDMETHOD( GetBandSiteInfo )( THIS_ BANDSITEINFO * ) PURE;
};
#endif

/* IModalWindow interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   IModalWindow
DECLARE_INTERFACE_( IModalWindow, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IModalWindow methods */
    STDMETHOD( Show )( THIS_ HWND ) PURE;
};
#endif

/* ICDBurnExt interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   ICDBurnExt
DECLARE_INTERFACE_( ICDBurnExt, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICDBurnExt methods */
    STDMETHOD( GetSupportedActionTypes )( THIS_ CDBE_ACTIONS * ) PURE;
};
#endif

/* IContextMenuSite interface */
#undef INTERFACE
#define INTERFACE   IContextMenuSite
DECLARE_INTERFACE_( IContextMenuSite, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IContextMenuSite methods */
    STDMETHOD( DoContextMenuPopup )( THIS_ IUnknown *, UINT, POINT ) PURE;
};

/* IEnumReadyCallback interface */
#undef INTERFACE
#define INTERFACE   IEnumReadyCallback
DECLARE_INTERFACE_( IEnumReadyCallback, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumReadyCallback methods */
    STDMETHOD( EnumReady )( THIS ) PURE;
};

/* IEnumerableView interface */
#undef INTERFACE
#define INTERFACE   IEnumerableView
DECLARE_INTERFACE_( IEnumerableView, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumerableView methods */
    STDMETHOD( SetEnumReadyCallback )( THIS_ IEnumReadyCallback * ) PURE;
    STDMETHOD( CreateEnumIDListFromContents )( THIS_ PCIDLIST_ABSOLUTE, DWORD, IEnumIDList ** ) PURE;
};

/* IInsertItem interface */
#if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IInsertItem
DECLARE_INTERFACE_( IInsertItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInsertItem methods */
    STDMETHOD( InsertItem )( THIS_ PCUIDLIST_RELATIVE ) PURE;
};
#endif

/* IMenuBand interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   IMenuBand
DECLARE_INTERFACE_( IMenuBand, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IMenuBand methods */
    STDMETHOD( IsMenuMessage )( THIS_ MSG * ) PURE;
    STDMETHOD( TranslateMenuMessage )( THIS_ MSG *, LRESULT * ) PURE;
};
#endif

/* IFolderBandPriv interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   IFolderBandPriv
DECLARE_INTERFACE_( IFolderBandPriv, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFolderBandPriv methods */
    STDMETHOD( SetCascade )( THIS_ BOOL ) PURE;
    STDMETHOD( SetAccelerators )( THIS_ BOOL ) PURE;
    STDMETHOD( SetNoIcons )( THIS_ BOOL ) PURE;
    STDMETHOD( SetNoText )( THIS_ BOOL ) PURE;
};
#endif

/* IRegTreeItem interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   IRegTreeItem
DECLARE_INTERFACE_( IRegTreeItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRegTreeItem methods */
    STDMETHOD( GetCheckState )( THIS_ BOOL * ) PURE;
    STDMETHOD( SetCheckState )( THIS_ BOOL ) PURE;
};
#endif

/* IImageRecompress interface */
#if (NTDDI_VERSION >= 0x05010000)
#undef INTERFACE
#define INTERFACE   IImageRecompress
DECLARE_INTERFACE_( IImageRecompress, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IImageRecompress methods */
    STDMETHOD( RecompressImage )( THIS_ IShellItem *, int, int, int, IStorage *, IStream ** ) PURE;
};
#endif

/* IDeskBar interface */
#if (_WIN32_IE >= 0x0600)
#undef INTERFACE
#define INTERFACE   IDeskBar
DECLARE_INTERFACE_( IDeskBar, IOleWindow ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IDeskBar methods */
    STDMETHOD( SetClient )( THIS_ IUnknown * ) PURE;
    STDMETHOD( GetClient )( THIS_ IUnknown ** ) PURE;
    STDMETHOD( OnPosRectChangeDB )( THIS_ RECT * ) PURE;
};
#endif

/* IMenuPopup interface */
#if (_WIN32_IE >= 0x0600)
#undef INTERFACE
#define INTERFACE   IMenuPopup
DECLARE_INTERFACE_( IMenuPopup, IDeskBar ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;
    
    /* IOleWindow methods */
    STDMETHOD( GetWindow )( THIS_ HWND * ) PURE;
    STDMETHOD( ContextSensitiveHelp )( THIS_ BOOL ) PURE;

    /* IDeskBar methods */
    STDMETHOD( SetClient )( THIS_ IUnknown * ) PURE;
    STDMETHOD( GetClient )( THIS_ IUnknown ** ) PURE;
    STDMETHOD( OnPosRectChangeDB )( THIS_ RECT * ) PURE;

    /* IMenuPopup methods */
    STDMETHOD( Popup )( THIS_ POINTL *, RECTL *, MF_POPUPFLAGS ) PURE;
    STDMETHOD( OnSelect )( THIS_ DWORD ) PURE;
    STDMETHOD( SetSubMenu )( THIS_ IMenuPopup *, BOOL ) PURE;
};
#endif

/* IFileIsInUse interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileIsInUse
DECLARE_INTERFACE_( IFileIsInUse, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileIsInUse methods */
    STDMETHOD( GetAppName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( GetUsage )( THIS_ FILE_USAGE_TYPE * ) PURE;
    STDMETHOD( GetCapabilities )( THIS_ DWORD * ) PURE;
    STDMETHOD( GetSwitchToHWND )( THIS_ HWND * ) PURE;
    STDMETHOD( CloseFile )( THIS ) PURE;
};
#endif

/* IFileDialogEvents interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileDialogEvents
DECLARE_INTERFACE_( IFileDialogEvents, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileDialogEvents methods */
    STDMETHOD( OnFileOk )( THIS_ IFileDialog * ) PURE;
    STDMETHOD( OnFolderChanging )( THIS_ IFileDialog *, IShellItem * ) PURE;
    STDMETHOD( OnFolderChange )( THIS_ IFileDialog * ) PURE;
    STDMETHOD( OnSelectionChange )( THIS_ IFileDialog * ) PURE;
    STDMETHOD( OnShareViolation )( THIS_ IFileDialog *, IShellItem *, FDE_SHAREVIOLATION_RESPONSE * ) PURE;
    STDMETHOD( OnTypeChange )( THIS_ IFileDialog * ) PURE;
    STDMETHOD( OnOverwrite )( THIS_ IFileDialog *, IShellItem *, FDE_OVERWRITE_RESPONSE * ) PURE;
};
#endif

/* IFileDialog interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileDialog
DECLARE_INTERFACE_( IFileDialog, IModalWindow ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IModalWindow methods */
    STDMETHOD( Show )( THIS_ HWND ) PURE;

    /* IFileDialog methods */
    STDMETHOD( SetFileTypes )( THIS_ UINT, const COMDLG_FILTERSPEC * ) PURE;
    STDMETHOD( SetFileTypeIndex )( THIS_ UINT ) PURE;
    STDMETHOD( GetFileTypeIndex )( THIS_ UINT * ) PURE;
    STDMETHOD( Advise )( THIS_ IFileDialogEvents *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetOptions )( THIS_ DWORD ) PURE;
    STDMETHOD( GetOptions )( THIS_ DWORD * ) PURE;
    STDMETHOD( SetDefaultFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetFolder )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( GetCurrentSelection )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( SetFileName )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetFileName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( SetTitle )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetOkButtonLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFileNameLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetResult )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( AddPlace )( THIS_ IShellItem *, FDAP ) PURE;
    STDMETHOD( SetDefaultExtension )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Close )( THIS_ HRESULT ) PURE;
    STDMETHOD( SetClientGuid )( THIS_ REFGUID ) PURE;
    STDMETHOD( ClearClientData )( THIS ) PURE;
    STDMETHOD( SetFilter )( THIS_ IShellItemFilter * ) PURE;
};
#endif

/* IFileSaveDialog interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileSaveDialog
DECLARE_INTERFACE_( IFileSaveDialog, IFileDialog ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IModalWindow methods */
    STDMETHOD( Show )( THIS_ HWND ) PURE;

    /* IFileDialog methods */
    STDMETHOD( SetFileTypes )( THIS_ UINT, const COMDLG_FILTERSPEC * ) PURE;
    STDMETHOD( SetFileTypeIndex )( THIS_ UINT ) PURE;
    STDMETHOD( GetFileTypeIndex )( THIS_ UINT * ) PURE;
    STDMETHOD( Advise )( THIS_ IFileDialogEvents *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetOptions )( THIS_ DWORD ) PURE;
    STDMETHOD( GetOptions )( THIS_ DWORD * ) PURE;
    STDMETHOD( SetDefaultFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetFolder )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( GetCurrentSelection )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( SetFileName )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetFileName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( SetTitle )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetOkButtonLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFileNameLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetResult )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( AddPlace )( THIS_ IShellItem *, FDAP ) PURE;
    STDMETHOD( SetDefaultExtension )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Close )( THIS_ HRESULT ) PURE;
    STDMETHOD( SetClientGuid )( THIS_ REFGUID ) PURE;
    STDMETHOD( ClearClientData )( THIS ) PURE;
    STDMETHOD( SetFilter )( THIS_ IShellItemFilter * ) PURE;

    /* IFileSaveDialog methods */
    STDMETHOD( SetSaveAsItem )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetProperties )( THIS_ IPropertyStore * ) PURE;
    STDMETHOD( SetCollectedProperties )( THIS_ IPropertyDescriptionList *, BOOL ) PURE;
    STDMETHOD( GetProperties )( THIS_ IPropertyStore ** ) PURE;
    STDMETHOD( ApplyProperties )( THIS_ IShellItem *, IPropertyStore *, HWND, IFileOperationProgressSink * ) PURE;
};
#endif

/* IFileOpenDialog interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileOpenDialog
DECLARE_INTERFACE_( IFileOpenDialog, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IModalWindow methods */
    STDMETHOD( Show )( THIS_ HWND ) PURE;

    /* IFileDialog methods */
    STDMETHOD( SetFileTypes )( THIS_ UINT, const COMDLG_FILTERSPEC * ) PURE;
    STDMETHOD( SetFileTypeIndex )( THIS_ UINT ) PURE;
    STDMETHOD( GetFileTypeIndex )( THIS_ UINT * ) PURE;
    STDMETHOD( Advise )( THIS_ IFileDialogEvents *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetOptions )( THIS_ DWORD ) PURE;
    STDMETHOD( GetOptions )( THIS_ DWORD * ) PURE;
    STDMETHOD( SetDefaultFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetFolder )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( GetCurrentSelection )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( SetFileName )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetFileName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( SetTitle )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetOkButtonLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFileNameLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetResult )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( AddPlace )( THIS_ IShellItem *, FDAP ) PURE;
    STDMETHOD( SetDefaultExtension )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Close )( THIS_ HRESULT ) PURE;
    STDMETHOD( SetClientGuid )( THIS_ REFGUID ) PURE;
    STDMETHOD( ClearClientData )( THIS ) PURE;
    STDMETHOD( SetFilter )( THIS_ IShellItemFilter * ) PURE;

    /* IFileOpenDialog methods */
    STDMETHOD( GetResults )( THIS_ IShellItemArray ** ) PURE;
    STDMETHOD( GetSelectedItems )( THIS_ IShellItemArray ** ) PURE;
};
#endif

/* IFileDialogCustomize interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileDialogCustomize
DECLARE_INTERFACE_( IFileDialogCustomize, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileDialogCustomize methods */
    STDMETHOD( EnableOpenDropDown )( THIS_ DWORD ) PURE;
    STDMETHOD( AddMenu )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( AddPushButton )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( AddComboBox )( THIS_ DWORD ) PURE;
    STDMETHOD( AddRadioButtonList )( THIS_ DWORD ) PURE;
    STDMETHOD( AddCheckButton )( THIS_ DWORD, LPCWSTR, BOOL ) PURE;
    STDMETHOD( AddEditBox )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( AddSeparator )( THIS_ DWORD ) PURE;
    STDMETHOD( AddText )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( SetControlLabel )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( GetControlLabel )( THIS_ DWORD, CDCONTROLSTATEF * ) PURE;
    STDMETHOD( SetControlState )( THIS_ DWORD, CDCONTROLSTATEF ) PURE;
    STDMETHOD( GetEditBoxText )( THIS_ DWORD, WCHAR ** ) PURE;
    STDMETHOD( SetEditBoxText )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( GetCheckButtonState )( THIS_ DWORD, BOOL * ) PURE;
    STDMETHOD( SetCheckButtonState )( THIS_ DWORD, BOOL ) PURE;
    STDMETHOD( AddControlItem )( THIS_ DWORD, DWORD, LPCWSTR ) PURE;
    STDMETHOD( RemoveControlItem )( THIS_ DWORD, DWORD ) PURE;
    STDMETHOD( RemoveAllControlItems )( THIS_ DWORD ) PURE;
    STDMETHOD( GetControlItemState )( THIS_ DWORD, DWORD, CDCONTROLSTATEF * ) PURE;
    STDMETHOD( SetControlItemState )( THIS_ DWORD, DWORD, CDCONTROLSTATEF ) PURE;
    STDMETHOD( GetSelectedControlItem )( THIS_ DWORD, DWORD * ) PURE;
    STDMETHOD( SetSelectedControlItem )( THIS_ DWORD, DWORD ) PURE;
    STDMETHOD( StartVisualGroup )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( EndVisualGroup )( THIS ) PURE;
    STDMETHOD( MakeProminent )( THIS_ DWORD ) PURE;
    STDMETHOD( SetControlItemText )( THIS_ DWORD, DWORD, LPCWSTR ) PURE;
};
#endif

/* IFileDialogControlEvents interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileDialogControlEvents
DECLARE_INTERFACE_( IFileDialogControlEvents, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileDialogControlEvents methods */
    STDMETHOD( OnItemSelected )( THIS_ IFileDialogCustomize *, DWORD, DWORD ) PURE;
    STDMETHOD( OnButtonClicked )( THIS_ IFileDialogCustomize *, DWORD ) PURE;
    STDMETHOD( OnCheckButtonToggled )( THIS_ IFileDialogCustomize *, DWORD, BOOL ) PURE;
    STDMETHOD( OnControlActivating )( THIS_ IFileDialogCustomize *, DWORD ) PURE;
};
#endif

/* IFileDialog2 interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IFileDialog2
DECLARE_INTERFACE_( IFileDialog2, IFileDialog ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IModalWindow methods */
    STDMETHOD( Show )( THIS_ HWND ) PURE;

    /* IFileDialog methods */
    STDMETHOD( SetFileTypes )( THIS_ UINT, const COMDLG_FILTERSPEC * ) PURE;
    STDMETHOD( SetFileTypeIndex )( THIS_ UINT ) PURE;
    STDMETHOD( GetFileTypeIndex )( THIS_ UINT * ) PURE;
    STDMETHOD( Advise )( THIS_ IFileDialogEvents *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( SetOptions )( THIS_ DWORD ) PURE;
    STDMETHOD( GetOptions )( THIS_ DWORD * ) PURE;
    STDMETHOD( SetDefaultFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetFolder )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( GetCurrentSelection )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( SetFileName )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetFileName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( SetTitle )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetOkButtonLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFileNameLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetResult )( THIS_ IShellItem ** ) PURE;
    STDMETHOD( AddPlace )( THIS_ IShellItem *, FDAP ) PURE;
    STDMETHOD( SetDefaultExtension )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Close )( THIS_ HRESULT ) PURE;
    STDMETHOD( SetClientGuid )( THIS_ REFGUID ) PURE;
    STDMETHOD( ClearClientData )( THIS ) PURE;
    STDMETHOD( SetFilter )( THIS_ IShellItemFilter * ) PURE;

    /* IFileDialog2 methods */
    STDMETHOD( SetCancelButtonLabel )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetNavigationRoot )( THIS_ IShellItem * ) PURE;
};
#endif

/* IApplicationAssociationRegistration interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IApplicationAssociationRegistration
DECLARE_INTERFACE_( IApplicationAssociationRegistration, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IApplicationAssociationRegistration methods */
    STDMETHOD( QueryCurrentDefault )( THIS_ LPCWSTR, ASSOCIATIONTYPE, ASSOCIATIONLEVEL, LPWSTR * ) PURE;
    STDMETHOD( QueryAppIsDefault )( THIS_ LPCWSTR, ASSOCIATIONTYPE, ASSOCIATIONLEVEL, LPCWSTR, BOOL * ) PURE;
    STDMETHOD( QueryAppIsDefaultAll )( THIS_ ASSOCIATIONLEVEL, LPCWSTR, BOOL * ) PURE;
    STDMETHOD( SetAppAsDefault )( THIS_ LPCWSTR, LPCWSTR, ASSOCIATIONTYPE ) PURE;
    STDMETHOD( SetAppAsDefaultAll )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( ClearUserAssociations )( THIS ) PURE;
};
#endif

/* IApplicationAssociationRegistrationUI interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IApplicationAssociationRegistrationUI
DECLARE_INTERFACE_( IApplicationAssociationRegistrationUI, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IApplicationAssociationRegistrationUI methods */
    STDMETHOD( LaunchAdvancedAssociationUI )( THIS_ LPCWSTR ) PURE;
};
#endif

/* IDelegateFolder interface */
#undef INTERFACE
#define INTERFACE   IDelegateFolder
DECLARE_INTERFACE_( IDelegateFolder, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDelegateFolder methods */
    STDMETHOD( SetItemAlloc )( THIS_ IMalloc * ) PURE;
};

/* IBrowserFrameOptions interface */
#if (_WIN32_IE >= 0x0600)
#undef INTERFACE
#define INTERFACE   IBrowserFrameOptions
DECLARE_INTERFACE_( IBrowserFrameOptions, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IBrowserFrameOptions methods */
    STDMETHOD( GetFrameOptions )( THIS_ BROWSERFRAMEOPTIONS, BROWSERFRAMEOPTIONS * ) PURE;
};
typedef IBrowserFrameOptions    *LPBROWSERFRAMEOPTIONS;
#endif

/* INewWindowManager interface */
#if (_WIN32_IE >= 0x0602)
#undef INTERFACE
#define INTERFACE   INewWindowManager
DECLARE_INTERFACE_( INewWindowManager, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INewWindowManager methods */
    STDMETHOD( EvaluateNewWindow )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR, LPCWSTR, BOOL, DWORD, DWORD ) PURE;
};
#endif

/* IAttachmentExecute interface */
#if (_WIN32_IE >= 0x0602)
#undef INTERFACE
#define INTERFACE   IAttachmentExecute
DECLARE_INTERFACE_( IAttachmentExecute, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IAttachmentExecute methods */
    STDMETHOD( SetClientTitle )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetClientGuid )( THIS_ REFGUID ) PURE;
    STDMETHOD( SetLocalPath )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetFileName )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetSource )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( SetReferrer )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( CheckPolicy )( THIS ) PURE;
    STDMETHOD( Prompt )( THIS_ HWND, ATTACHMENT_PROMPT, ATTACHMENT_ACTION * ) PURE;
    STDMETHOD( Save )( THIS ) PURE;
    STDMETHOD( Execute )( THIS_ HWND, LPCWSTR, HANDLE * ) PURE;
    STDMETHOD( SaveWithUI )( THIS_ HWND ) PURE;
    STDMETHOD( ClearClientState )( THIS ) PURE;
};
#endif

/* IShellMenuCallback interface */
#if (_WIN32_IE >= 0x0600)
#undef INTERFACE
#define INTERFACE   IShellMenuCallback
DECLARE_INTERFACE_( IShellMenuCallback, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellMenuCallback methods */
    STDMETHOD( CallbackSM )( THIS_ LPSMDATA, UINT, WPARAM, LPARAM ) PURE;
};
#endif

/* IShellMenu interface */
#if (_WIN32_IE >= 0x0600)
#undef INTERFACE
#define INTERFACE   IShellMenu
DECLARE_INTERFACE_( IShellMenu, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellMenu methods */
    STDMETHOD( Initialize )( THIS_ IShellMenuCallback *, UINT, UINT, DWORD ) PURE;
    STDMETHOD( GetMenuInfo )( THIS_ IShellMenuCallback **, UINT *, UINT *, DWORD * ) PURE;
    STDMETHOD( SetShellFolder )( THIS_ IShellFolder *, PCIDLIST_ABSOLUTE, HKEY, DWORD ) PURE;
    STDMETHOD( GetShellFolder )( THIS_ DWORD *, PIDLIST_ABSOLUTE *, REFIID, void ** ) PURE;
    STDMETHOD( SetMenu )( THIS_ HMENU, HWND, DWORD ) PURE;
    STDMETHOD( GetMenu )( THIS_ HMENU *, HWND *, DWORD * ) PURE;
    STDMETHOD( InvalidateItem )( THIS_ LPSMDATA, DWORD ) PURE;
    STDMETHOD( GetState )( THIS_ LPSMDATA ) PURE;
    STDMETHOD( SetMenuToolbar )( THIS_ IUnknown *, DWORD ) PURE;
};
#endif

/* IShellRunDll interface */
#undef INTERFACE
#define INTERFACE   IShellRunDll
DECLARE_INTERFACE_( IShellRunDll, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellRunDll methods */
    STDMETHOD( Run )( THIS_ LPCWSTR ) PURE;
};

/* IKnownFolder interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IKnownFolder
DECLARE_INTERFACE_( IKnownFolder, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IKnownFolder methods */
    STDMETHOD( GetId )( THIS_ KNOWNFOLDERID * ) PURE;
    STDMETHOD( GetCategory )( THIS_ KF_CATEGORY * ) PURE;
    STDMETHOD( GetShellItem )( THIS_ DWORD, REFIID, void ** ) PURE;
    STDMETHOD( GetPath )( THIS_ DWORD, LPWSTR * ) PURE;
    STDMETHOD( SetPath )( THIS_ DWORD, LPCWSTR ) PURE;
    STDMETHOD( GetIDList )( THIS_ DWORD, PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetFolderType )( THIS_ FOLDERTYPEID * ) PURE;
    STDMETHOD( GetRedirectionCapabilities )( THIS_ KF_REDIRECTION_CAPABILITIES * ) PURE;
    STDMETHOD( GetFolderDefinition )( THIS_ KNOWNFOLDER_DEFINITION * ) PURE;
};
#endif

/* IKnownFolderManager interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IKnownFolderManager
DECLARE_INTERFACE_( IKnownFolderManager, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IKnownFolderManager methods */
    STDMETHOD( FolderIdFromCsidl )( THIS_ int, KNOWNFOLDERID * ) PURE;
    STDMETHOD( FolderIdToCsidl )( THIS_ REFKNOWNFOLDERID, int * ) PURE;
    STDMETHOD( GetFolderIds )( THIS_ KNOWNFOLDERID **, UINT * ) PURE;
    STDMETHOD( GetFolder )( THIS_ REFKNOWNFOLDERID, IKnownFolder ** ) PURE;
    STDMETHOD( GetFolderByName )( THIS_ LPCWSTR, IKnownFolder ** ) PURE;
    STDMETHOD( RegisterFolder )( THIS_ REFKNOWNFOLDERID, const KNOWNFOLDER_DEFINITION * ) PURE;
    STDMETHOD( UnregisterFolder )( THIS_ REFKNOWNFOLDERID ) PURE;
    STDMETHOD( FindFolderFromPath )( THIS_ LPCWSTR, FFFP_MODE, IKnownFolder ** ) PURE;
    STDMETHOD( FindFolderFromIDList )( THIS_ PCIDLIST_ABSOLUTE, IKnownFolder ** ) PURE;
    STDMETHOD( Redirect )( THIS_ REFKNOWNFOLDERID, HWND, KF_REDIRECT_FLAGS, LPCWSTR, UINT, const KNOWNFOLDERID *, LPWSTR * ) PURE;
};
#endif

/* ISharingConfigurationManager interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   ISharingConfigurationManager
DECLARE_INTERFACE_( ISharingConfigurationManager, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ISharingConfigurationManager methods */
    STDMETHOD( CreateShare )( THIS_ DEF_SHARE_ID, SHARE_ROLE ) PURE;
    STDMETHOD( DeleteShare )( THIS_ DEF_SHARE_ID ) PURE;
    STDMETHOD( ShareExists )( THIS_ DEF_SHARE_ID ) PURE;
    STDMETHOD( GetSharePermissions )( THIS_ DEF_SHARE_ID, SHARE_ROLE * ) PURE;
    STDMETHOD( SharePrinters )( THIS ) PURE;
    STDMETHOD( StopSharingPrinters )( THIS ) PURE;
    STDMETHOD( ArePrintersShared )( THIS ) PURE;
};
#endif

/* IPreviousVersionsInfo interface */
#undef INTERFACE
#define INTERFACE   IPreviousVersionsInfo
DECLARE_INTERFACE_( IPreviousVersionsInfo, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IPreviousVersionsInfo methods */
    STDMETHOD( AreSnapshotsAvailable )( THIS_ LPCWSTR, BOOL, BOOL * ) PURE;
};

/* IRelatedItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IRelatedItem
DECLARE_INTERFACE_( IRelatedItem, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IIdentityName interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IIdentityName
DECLARE_INTERFACE_( IIdentityName, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IDelegateItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IDelegateItem
DECLARE_INTERFACE_( IDelegateItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* ICurrentItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   ICurrentItem
DECLARE_INTERFACE_( ICurrentItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* ITransferMediumItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   ITransferMediumItem
DECLARE_INTERFACE_( ITransferMediumItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IUseToBrowseItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IUseToBrowseItem
DECLARE_INTERFACE_( IUseToBrowseItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IDisplayItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IDisplayItem
DECLARE_INTERFACE_( IDisplayItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IViewStateIdentityItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IViewStateIdentityItem
DECLARE_INTERFACE_( IViewStateIdentityItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IPreviewItem interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IPreviewItem
DECLARE_INTERFACE_( IPreviewItem, IRelatedItem ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IRelatedItem methods */
    STDMETHOD( GetItemIDList )( THIS_ PIDLIST_ABSOLUTE * ) PURE;
    STDMETHOD( GetItem )( THIS_ IShellItem ** ) PURE;
};
#endif

/* IDestinationStreamFactory interface */
#undef INTERFACE
#define INTERFACE   IDestinationStreamFactory
DECLARE_INTERFACE_( IDestinationStreamFactory, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDestinationStreamFactory methods */
    STDMETHOD( GetDestinationStream )( THIS_ IStream ** ) PURE;
};

/* INewMenuClient interface */
#undef INTERFACE
#define INTERFACE   INewMenuClient
DECLARE_INTERFACE_( INewMenuClient, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INewMenuClient methods */
    STDMETHOD( IncludeItems )( THIS_ NMCII_FLAGS * ) PURE;
    STDMETHOD( SelectAndEditItem )( THIS_ PCIDLIST_ABSOLUTE, NMCSAEI_FLAGS ) PURE;
};

/* IInitializeWithBindCtx interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IInitializeWithBindCtx
DECLARE_INTERFACE_( IInitializeWithBindCtx, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInitializeWithBindCtx methods */
    STDMETHOD( Initialize )( THIS_ IBindCtx * ) PURE;
};
#endif

/* IShellItemFilter interface */
#if (_WIN32_IE >= 0x0700)
#undef INTERFACE
#define INTERFACE   IShellItemFilter
DECLARE_INTERFACE_( IShellItemFilter, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellItemFilter methods */
    STDMETHOD( IncludeItem )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetEnumFlagsForItem )( THIS_ IShellItem *, SHCONTF * ) PURE;
};
#endif

/* INameSpaceTreeControl interface */
#undef INTERFACE
#define INTERFACE   INameSpaceTreeControl
DECLARE_INTERFACE_( INameSpaceTreeControl, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeControl methods */
    STDMETHOD( Initialize )( THIS_ HWND, RECT *, NSTCSTYLE ) PURE;
    STDMETHOD( TreeAdvise )( THIS_ IUnknown *, DWORD * ) PURE;
    STDMETHOD( TreeUnadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( AppendRoot )( THIS_ IShellItem *, SHCONTF, NSTCROOTSTYLE, IShellItemFilter * ) PURE;
    STDMETHOD( InsertRoot )( THIS_ int, IShellItem *, SHCONTF, NSTCROOTSTYLE, IShellItemFilter * ) PURE;
    STDMETHOD( RemoveRoot )( THIS_ IShellItem * ) PURE;
    STDMETHOD( RemoveAllRoots )( THIS ) PURE;
    STDMETHOD( GetRootItems )( THIS_ IShellItemArray ** ) PURE;
    STDMETHOD( SetItemState )( THIS_ IShellItem *, NSTCITEMSTATE, NSTCITEMSTATE ) PURE;
    STDMETHOD( GetItemState )( THIS_ IShellItem *, NSTCITEMSTATE, NSTCITEMSTATE * ) PURE;
    STDMETHOD( GetSelectedItems )( THIS_ IShellItemArray ** ) PURE;
    STDMETHOD( GetItemCustomState )( THIS_ IShellItem *, int * ) PURE;
    STDMETHOD( SetItemCustomState )( THIS_ IShellItem *, int ) PURE;
    STDMETHOD( EnsureItemVisible )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetTheme )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetNextItem )( THIS_ IShellItem *, NSTCGNI, IShellItem ** ) PURE;
    STDMETHOD( HitTest )( THIS_ POINT *, IShellItem ** ) PURE;
    STDMETHOD( GetItemRect )( THIS_ IShellItem *, RECT * ) PURE;
    STDMETHOD( CollapseAll )( THIS ) PURE;
};

/* INameSpaceTreeControl2 interface */
#undef INTERFACE
#define INTERFACE   INameSpaceTreeControl2
DECLARE_INTERFACE_( INameSpaceTreeControl2, INameSpaceTreeControl ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeControl methods */
    STDMETHOD( Initialize )( THIS_ HWND, RECT *, NSTCSTYLE ) PURE;
    STDMETHOD( TreeAdvise )( THIS_ IUnknown *, DWORD * ) PURE;
    STDMETHOD( TreeUnadvise )( THIS_ DWORD ) PURE;
    STDMETHOD( AppendRoot )( THIS_ IShellItem *, SHCONTF, NSTCROOTSTYLE, IShellItemFilter * ) PURE;
    STDMETHOD( InsertRoot )( THIS_ int, IShellItem *, SHCONTF, NSTCROOTSTYLE, IShellItemFilter * ) PURE;
    STDMETHOD( RemoveRoot )( THIS_ IShellItem * ) PURE;
    STDMETHOD( RemoveAllRoots )( THIS ) PURE;
    STDMETHOD( GetRootItems )( THIS_ IShellItemArray ** ) PURE;
    STDMETHOD( SetItemState )( THIS_ IShellItem *, NSTCITEMSTATE, NSTCITEMSTATE ) PURE;
    STDMETHOD( GetItemState )( THIS_ IShellItem *, NSTCITEMSTATE, NSTCITEMSTATE * ) PURE;
    STDMETHOD( GetSelectedItems )( THIS_ IShellItemArray ** ) PURE;
    STDMETHOD( GetItemCustomState )( THIS_ IShellItem *, int * ) PURE;
    STDMETHOD( SetItemCustomState )( THIS_ IShellItem *, int ) PURE;
    STDMETHOD( EnsureItemVisible )( THIS_ IShellItem * ) PURE;
    STDMETHOD( SetTheme )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetNextItem )( THIS_ IShellItem *, NSTCGNI, IShellItem ** ) PURE;
    STDMETHOD( HitTest )( THIS_ POINT *, IShellItem ** ) PURE;
    STDMETHOD( GetItemRect )( THIS_ IShellItem *, RECT * ) PURE;
    STDMETHOD( CollapseAll )( THIS ) PURE;

    /* INameSpaceTreeControl2 methods */
    STDMETHOD( SetControlStyle )( THIS_ NSTCSTYLE, NSTCSTYLE ) PURE;
    STDMETHOD( GetControlStyle )( THIS_ NSTCSTYLE, NSTCSTYLE * ) PURE;
    STDMETHOD( SetControlStyle2 )( THIS_ NSTCSTYLE2, NSTCSTYLE2 ) PURE;
    STDMETHOD( GetControlStyle2 )( THIS_ NSTCSTYLE2, NSTCSTYLE2 * ) PURE;
};

/* INameSpaceTreeControlEvents interface */
#undef INTERFACE
#define INTERFACE   INameSpaceTreeControlEvents
DECLARE_INTERFACE_( INameSpaceTreeControlEvents, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeControlEvents methods */
    STDMETHOD( OnItemClick )( THIS_ IShellItem *, NSTCEHITTEST, NSTCECLICKTYPE ) PURE;
    STDMETHOD( OnPropertyItemCommit )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnItemStateChanging )( THIS_ IShellItem *, NSTCITEMSTATE, NSTCITEMSTATE ) PURE;
    STDMETHOD( OnItemStateChanged )( THIS_ IShellItem *, NSTCITEMSTATE, NSTCITEMSTATE ) PURE;
    STDMETHOD( OnSelectionChanged )( THIS_ IShellItemArray * ) PURE;
    STDMETHOD( OnKeyboardInput )( THIS_ UINT, WPARAM, LPARAM ) PURE;
    STDMETHOD( OnBeforeExpand )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnAfterExpand )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnBeginLabelEdit )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnEndLabelEdit )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnGetToolTip )( THIS_ IShellItem *, LPWSTR, int ) PURE;
    STDMETHOD( OnBeforeItemDelete )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnItemAdded )( THIS_ IShellItem *, BOOL ) PURE;
    STDMETHOD( OnItemDeleted )( THIS_ IShellItem *, BOOL ) PURE;
    STDMETHOD( OnBeforeContextMenu )( THIS_ IShellItem *, REFIID, void ** ) PURE;
    STDMETHOD( OnAfterContextMenu )( THIS_ IShellItem *, IContextMenu *, REFIID, void ** ) PURE;
    STDMETHOD( OnBeforeStateImageChange )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnGetDefaultIconIndex )( THIS_ IShellItem *, int *, int * ) PURE;
};

/* INameSpaceTreeControlDropHandler interface */
#undef INTERFACE
#define INTERFACE   INameSpaceTreeControlDropHandler
DECLARE_INTERFACE_( INameSpaceTreeControlDropHandler, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeControlDropHandler methods */
    STDMETHOD( OnDragEnter )( THIS_ IShellItem *, IShellItemArray *, BOOL, DWORD, DWORD * ) PURE;
    STDMETHOD( OnDragOver )( THIS_ IShellItem *, IShellItemArray *, DWORD, DWORD * ) PURE;
    STDMETHOD( OnDragPosition )( THIS_ IShellItem *, IShellItemArray *, int, int ) PURE;
    STDMETHOD( OnDrop )( THIS_ IShellItem *, IShellItemArray *, int, DWORD, DWORD * ) PURE;
    STDMETHOD( OnDropPosition )( THIS_ IShellItem *, IShellItemArray *, int, int ) PURE;
    STDMETHOD( OnDragLeave )( THIS_ IShellItem * ) PURE;
};

/* INameSpaceTreeAccessible interface */
#undef INTERFACE
#define INTERFACE   INameSpaceTreeAccessible
DECLARE_INTERFACE_( INameSpaceTreeAccessible, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeAccessible methods */
    STDMETHOD( OnGetDefaultAccessibilityAction )( THIS_ IShellItem *, BSTR * ) PURE;
    STDMETHOD( OnDoDefaultAccessibilityAction )( THIS_ IShellItem * ) PURE;
    STDMETHOD( OnGetAccessibilityRole )( THIS_ IShellItem *, VARIANT * ) PURE;
};

/* INameSpaceTreeControlCustomDraw interface */
#undef INTERFACE
#define INTERFACE   INameSpaceTreeControlCustomDraw
DECLARE_INTERFACE_( INameSpaceTreeControlCustomDraw, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeControlCustomDraw methods */
    STDMETHOD( PrePaint )( THIS_ HDC, RECT *, LRESULT * ) PURE;
    STDMETHOD( PostPaint )( THIS_ HDC, RECT * ) PURE;
    STDMETHOD( ItemPrePaint )( THIS_ HDC, RECT *, NSTCCUSTOMDRAW *, COLORREF *, COLORREF *, LRESULT * ) PURE;
    STDMETHOD( ItemPostPaint )( THIS_ HDC, RECT *, NSTCCUSTOMDRAW * ) PURE;
};

/* INameSpaceTreeControlFolderCapabilities interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   INameSpaceTreeControlFolderCapabilities
DECLARE_INTERFACE_( INameSpaceTreeControlFolderCapabilities, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* INameSpaceTreeControlFolderCapabilities methods */
    STDMETHOD( GetFolderCapabilities )( THIS_ NSTCFOLDERCAPABILITIES, NSTCFOLDERCAPABILITIES * ) PURE;
};
#endif

/* IPreviewHandler interface */
#undef INTERFACE
#define INTERFACE   IPreviewHandler
DECLARE_INTERFACE_( IPreviewHandler, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IPreviewHandler methods */
    STDMETHOD( SetWindow )( THIS_ HWND, const RECT * ) PURE;
    STDMETHOD( SetRect )( THIS_ const RECT * ) PURE;
    STDMETHOD( DoPreview )( THIS ) PURE;
    STDMETHOD( Unload )( THIS ) PURE;
    STDMETHOD( SetFocus )( THIS ) PURE;
    STDMETHOD( QueryFocus )( THIS_ HWND * ) PURE;
    STDMETHOD( TranslateAccelerator )( THIS_ MSG * ) PURE;
};

/* IPreviewHandlerFrame interface */
#undef INTERFACE
#define INTERFACE   IPreviewHandlerFrame
DECLARE_INTERFACE_( IPreviewHandlerFrame, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IPreviewHandlerFrame methods */
    STDMETHOD( GetWindowContext )( THIS_ PREVIEWHANDLERFRAMEINFO * ) PURE;
    STDMETHOD( TranslateAccelerator )( THIS_ MSG * ) PURE;
};

/* ITrayDeskBand interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   ITrayDeskBand
DECLARE_INTERFACE_( ITrayDeskBand, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ITrayDeskBand methods */
    STDMETHOD( ShowDeskBand )( THIS_ REFCLSID ) PURE;
    STDMETHOD( HideDeskBand )( THIS_ REFCLSID ) PURE;
    STDMETHOD( IsDeskBandShown )( THIS_ REFCLSID ) PURE;
    STDMETHOD( DeskBandRegistrationChanged )( THIS ) PURE;
};
#endif

/* IBandHost interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IBandHost
DECLARE_INTERFACE_( IBandHost, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IBandHost methods */
    STDMETHOD( CreateBand )( THIS_ REFCLSID, BOOL, BOOL, REFIID, void ** ) PURE;
    STDMETHOD( SetBandAvailability )( THIS_ REFCLSID, BOOL ) PURE;
    STDMETHOD( DestroyBand )( THIS_ REFCLSID ) PURE;
};
#endif

/* IExplorerPaneVisibility interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IExplorerPaneVisibility
DECLARE_INTERFACE_( IExplorerPaneVisibility, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExplorerPaneVisibility methods */
    STDMETHOD( GetPaneState )( THIS_ REFEXPLORERPANE, EXPLORERPANESTATE * ) PURE;
};
#endif

/* IContextMenuCB interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IContextMenuCB
DECLARE_INTERFACE_( IContextMenuCB, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IContextMenuCB methods */
    STDMETHOD( CallBack )( THIS_ IShellFolder *, HWND, IDataObject *, UINT, WPARAM, LPARAM ) PURE;
};
#endif

/* IDefaultExtractIconInit interface */
#undef INTERFACE
#define INTERFACE   IDefaultExtractIconInit
DECLARE_INTERFACE_( IDefaultExtractIconInit, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDefaultExtractIconInit methods */
    STDMETHOD( SetFlags )( THIS_ UINT ) PURE;
    STDMETHOD( SetKey )( THIS_ HKEY ) PURE;
    STDMETHOD( SetNormalIcon )( THIS_ LPCWSTR, int ) PURE;
    STDMETHOD( SetOpenIcon )( THIS_ LPCWSTR, int ) PURE;
    STDMETHOD( SetShortcutIcon )( THIS_ LPCWSTR, int ) PURE;
    STDMETHOD( SetDefaultIcon )( THIS_ LPCWSTR, int ) PURE;
};

/* IExplorerCommand interface */
#undef INTERFACE
#define INTERFACE   IExplorerCommand
DECLARE_INTERFACE_( IExplorerCommand, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExplorerCommand methods */
    STDMETHOD( GetTitle )( THIS_ IShellItemArray *, LPWSTR * ) PURE;
    STDMETHOD( GetIcon )( THIS_ IShellItemArray *, LPWSTR * ) PURE;
    STDMETHOD( GetToolTip )( THIS_ IShellItemArray *, LPWSTR * ) PURE;
    STDMETHOD( GetCanonicalName )( THIS_ GUID * ) PURE;
    STDMETHOD( GetState )( THIS_ IShellItemArray *, BOOL, EXPCMDSTATE * ) PURE;
    STDMETHOD( Invoke )( THIS_ IShellItemArray *, IBindCtx * ) PURE;
    STDMETHOD( GetFlags )( THIS_ EXPCMDFLAGS * ) PURE;
    STDMETHOD( EnumSubCommands )( THIS_ IEnumExplorerCommand ** ) PURE;
};

/* IExplorerCommandState interface */
#undef INTERFACE
#define INTERFACE   IExplorerCommandState
DECLARE_INTERFACE_( IExplorerCommandState, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExplorerCommandState methods */
    STDMETHOD( GetState )( THIS_ IShellItemArray *, BOOL, EXPCMDSTATE * ) PURE;
};

/* IInitializeCommand interface */
#undef INTERFACE
#define INTERFACE   IInitializeCommand
DECLARE_INTERFACE_( IInitializeCommand, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInitializeCommand methods */
    STDMETHOD( Initialize )( THIS_ LPCWSTR, IPropertyBag * ) PURE;
};

/* IEnumExplorerCommand interface */
#undef INTERFACE
#define INTERFACE   IEnumExplorerCommand
DECLARE_INTERFACE_( IEnumExplorerCommand, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumExplorerCommand methods */
    STDMETHOD( Next )( THIS_ ULONG, IExplorerCommand **, ULONG * ) PURE;
    STDMETHOD( Skip )( THIS_ ULONG ) PURE;
    STDMETHOD( Reset )( THIS ) PURE;
    STDMETHOD( Clone )( THIS_ IEnumExplorerCommand ** ) PURE;
};

/* IExplorerCommandProvider interface */
#undef INTERFACE
#define INTERFACE   IExplorerCommandProvider
DECLARE_INTERFACE_( IExplorerCommandProvider, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IExplorerCommandProvider methods */
    STDMETHOD( GetCommands )( THIS_ IUnknown *, REFIID, void ** ) PURE;
    STDMETHOD( GetCommand )( THIS_ REFGUID, REFIID, void ** ) PURE;
};

/* IMarkupCallback interface (FOR COMPATIBILITY ONLY - NO LONGER IN MS HEADERS) */
#undef INTERFACE
#define INTERFACE   IMarkupCallback
DECLARE_INTERFACE_( IMarkupCallback, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IMarkupCallback methods */
    STDMETHOD( GetState )( THIS_ DWORD, UINT ) PURE;
    STDMETHOD( Notify )( THIS_ DWORD, int, int ) PURE;
    STDMETHOD( InvalidateRect )( THIS_ DWORD, const RECT * ) PURE;
    STDMETHOD( OnCustomDraw )( THIS_ DWORD, HDC, const RECT *, DWORD, int, UINT, LRESULT * ) PURE;
    STDMETHOD( CustomDrawText )( THIS_ HDC, LPCWSTR, int, RECT *, UINT, BOOL ) PURE;
};

/* IControlMarkup interface (FOR COMPATIBILITY ONLY - NO LONGER IN MS HEADERS) */
#undef INTERFACE
#define INTERFACE   IControlMarkup
DECLARE_INTERFACE_( IControlMarkup, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IControlMarkup methods */
    STDMETHOD( SetCallback )( THIS_ IUnknown * ) PURE;
    STDMETHOD( GetCallback )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( SetId )( THIS_ DWORD ) PURE;
    STDMETHOD( GetId )( THIS_ DWORD * ) PURE;
    STDMETHOD( SetFonts )( THIS_ HFONT, HFONT ) PURE;
    STDMETHOD( GetFonts )( THIS_ HFONT *, HFONT * ) PURE;
    STDMETHOD( SetText )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetText )( THIS_ BOOL, LPWSTR, DWORD * ) PURE;
    STDMETHOD( SetLinkText )( THIS_ int, UINT, LPCWSTR ) PURE;
    STDMETHOD( GetLinkText )( THIS_ int, UINT, LPWSTR, DWORD * ) PURE;
    STDMETHOD( SetRenderFlags )( THIS_ UINT ) PURE;
    STDMETHOD( GetRenderFlags )( THIS_ UINT *, HTHEME *, int *, int *, int * ) PURE;
    STDMETHOD( SetThemeRenderFlags )( THIS_ UINT, HTHEME, int, int, int ) PURE;
    STDMETHOD( GetState )( THIS_ int, UINT, UINT * ) PURE;
    STDMETHOD( SetState )( THIS_ int, UINT, UINT ) PURE;
    STDMETHOD( DrawText )( THIS_ HDC, LPCRECT ) PURE;
    STDMETHOD( SetLinkCursor )( THIS ) PURE;
    STDMETHOD( CalcIdealSize )( THIS_ HDC, UINT, RECT * ) PURE;
    STDMETHOD( SetFocus )( THIS ) PURE;
    STDMETHOD( KillFocus )( THIS ) PURE;
    STDMETHOD( IsTabbable )( THIS ) PURE;
    STDMETHOD( OnButtonDown )( THIS_ POINT ) PURE;
    STDMETHOD( OnButtonUp )( THIS_ POINT ) PURE;
    STDMETHOD( OnKeyDown )( THIS_ UINT ) PURE;
    STDMETHOD( HitTest )( THIS_ POINT, int * ) PURE;
    STDMETHOD( GetLinkRect )( THIS_ int, RECT * ) PURE;
    STDMETHOD( GetControlRect )( THIS_ RECT * ) PURE;
    STDMETHOD( GetLinkCount )( THIS_ UINT * ) PURE;
};

/* IInitializeNetworkFolder interface */
#undef INTERFACE
#define INTERFACE   IInitializeNetworkFolder
DECLARE_INTERFACE_( IInitializeNetworkFolder, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInitializeNetworkFolder methods */
    STDMETHOD( Initialize )( THIS_ PCIDLIST_ABSOLUTE, PCIDLIST_ABSOLUTE, UINT, LPCWSTR, LPCWSTR ) PURE;
};

/* IOpenControlPanel interface */
#undef INTERFACE
#define INTERFACE   IOpenControlPanel
DECLARE_INTERFACE_( IOpenControlPanel, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IOpenControlPanel methods */
    STDMETHOD( Open )( THIS_ LPCWSTR, LPCWSTR, IUnknown * ) PURE;
    STDMETHOD( GetPath )( THIS_ LPCWSTR, LPWSTR, UINT ) PURE;
    STDMETHOD( GetCurrentView )( THIS_ CPVIEW * ) PURE;
};

/* ISystemCPLUpdate interface (FOR COMPATIBILITY ONLY - NO LONGER IN MS HEADERS) */
#undef INTERFACE
#define INTERFACE   ISystemCPLUpdate
DECLARE_INTERFACE_( ISystemCPLUpdate, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ISystemCPLUpdate methods */
    STDMETHOD( UpdateSystemInfo )( THIS_ LPCWSTR, LPCWSTR, LPCWSTR, DWORD ) PURE;
    STDMETHOD( UpdateLicensingInfo )( THIS_ DWORD, DWORD, DWORD, BOOL ) PURE;
    STDMETHOD( UpdateRatingsInfo )( THIS_ LPCWSTR, LPCWSTR, HBITMAP, USHORT ) PURE;
    STDMETHOD( UpdateComputerInfo )( THIS ) PURE;
};

/* IComputerInfoAdvise interface (FOR COMPATIBILITY ONLY - NO LONGER IN MS HEADERS) */
#undef INTERFACE
#define INTERFACE   IComputerInfoAdvise
DECLARE_INTERFACE_( IComputerInfoAdvise, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IComputerInfoAdvise methods */
    STDMETHOD( Advise )( THIS_ ISystemCPLUpdate *, DWORD * ) PURE;
    STDMETHOD( Unadvise )( THIS_ DWORD ) PURE;
};

/* IComputerInfoChangeNotify interface */
#undef INTERFACE
#define INTERFACE   IComputerInfoChangeNotify
DECLARE_INTERFACE_( IComputerInfoChangeNotify, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IComputerInfoChangeNotify methods */
    STDMETHOD( ComputerInfoChanged )( THIS ) PURE;
};

/* IFileSystemBindData interface */
#undef INTERFACE
#define INTERFACE   IFileSystemBindData
DECLARE_INTERFACE_( IFileSystemBindData, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileSystemBindData methods */
    STDMETHOD( SetFindData )( THIS_ const WIN32_FIND_DATAW * ) PURE;
    STDMETHOD( GetFindData )( THIS_ WIN32_FIND_DATAW * ) PURE;
};

/* IFileSystemBindData2 interface */
#undef INTERFACE
#define INTERFACE   IFileSystemBindData2
DECLARE_INTERFACE_( IFileSystemBindData2, IFileSystemBindData ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IFileSystemBindData methods */
    STDMETHOD( SetFindData )( THIS_ const WIN32_FIND_DATAW * ) PURE;
    STDMETHOD( GetFindData )( THIS_ WIN32_FIND_DATAW * ) PURE;

    /* IFileSystemBindData2 methods */
    STDMETHOD( SetFileID )( THIS_ LARGE_INTEGER ) PURE;
    STDMETHOD( GetFileID )( THIS_ LARGE_INTEGER * ) PURE;
    STDMETHOD( SetJunctionCLSID )( THIS_ REFCLSID ) PURE;
    STDMETHOD( GetJunctionCLSID )( THIS_ CLSID * ) PURE;
};

/* ICustomDestinationList interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   ICustomDestinationList
DECLARE_INTERFACE_( ICustomDestinationList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* ICustomDestinationList methods */
    STDMETHOD( SetAppID )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( BeginList )( THIS_ UINT *, REFIID, void ** ) PURE;
    STDMETHOD( AppendCategory )( THIS_ LPCWSTR, IObjectArray * ) PURE;
    STDMETHOD( AppendKnownCategory )( THIS_ KNOWNDESTCATEGORY ) PURE;
    STDMETHOD( AddUserTasks )( THIS_ IObjectArray * ) PURE;
    STDMETHOD( CommitList )( THIS ) PURE;
    STDMETHOD( GetRemovedDestinations )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD( DeleteList )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( AbortList )( THIS ) PURE;
};
#endif

/* IApplicationDestinations interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IApplicationDestinations
DECLARE_INTERFACE_( IApplicationDestinations, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IApplicationDestinations methods */
    STDMETHOD( SetAppID )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( RemoveDestination )( THIS_ IUnknown * ) PURE;
    STDMETHOD( RemoveAllDestinations )( THIS ) PURE;
};
#endif

/* IApplicationDocumentLists interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IApplicationDocumentLists
DECLARE_INTERFACE_( IApplicationDocumentLists, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IApplicationDocumentLists methods */
    STDMETHOD( SetAppID )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetList )( THIS_ APPDOCLISTTYPE, UINT, REFIID, void ** ) PURE;
};
#endif

/* IObjectWithAppUserModelID interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IObjectWithAppUserModelID
DECLARE_INTERFACE_( IObjectWithAppUserModelID, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectWithAppUserModelID methods */
    STDMETHOD( SetAppID )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetAppID )( THIS_ LPWSTR * ) PURE;
};
#endif

/* IObjectWithProgID interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IObjectWithProgID
DECLARE_INTERFACE_( IObjectWithProgID, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IObjectWithProgID methods */
    STDMETHOD( SetProgID )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( GetProgID )( THIS_ LPWSTR * ) PURE;
};
#endif

/* IUpdateIDList interface */
#if (NTDDI_VERSION >= 0x06010000)
#undef INTERFACE
#define INTERFACE   IUpdateIDList
DECLARE_INTERFACE_( IUpdateIDList, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IUpdateIDList methods */
    STDMETHOD( Update )( THIS_ IBindCtx *, PCUITEMID_CHILD, PITEMID_CHILD * ) PURE;
};
#endif

/* IDesktopGadget interface */
#undef INTERFACE
#define INTERFACE   IDesktopGadget
DECLARE_INTERFACE_( IDesktopGadget, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IDesktopGadget methods */
    STDMETHOD( RunGadget )( THIS_ LPCWSTR ) PURE;
};

/* IHomeGroup interface */
#undef INTERFACE
#define INTERFACE   IHomeGroup
DECLARE_INTERFACE_( IHomeGroup, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IHomeGroup methods */
    STDMETHOD( IsMember )( THIS_ BOOL * ) PURE;
    STDMETHOD( ShowSharingWizard )( THIS_ HWND, HOMEGROUPSHARINGCHOICES * ) PURE;
};

/* IInitializeWithPropertyStore interface */
#undef INTERFACE
#define INTERFACE   IInitializeWithPropertyStore
DECLARE_INTERFACE_( IInitializeWithPropertyStore, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IInitializeWithPropertyStore methods */
    STDMETHOD( Initialize )( THIS_ IPropertyStore * ) PURE;
};

/* IOpenSearchSource interface */
#undef INTERFACE
#define INTERFACE   IOpenSearchSource
DECLARE_INTERFACE_( IOpenSearchSource, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IOpenSearchSource methods */
    STDMETHOD( GetResults )( THIS_ HWND, LPCWSTR, DWORD, DWORD, REFIID, void ** ) PURE;
};

/* IShellLibrary interface */
#undef INTERFACE
#define INTERFACE   IShellLibrary
DECLARE_INTERFACE_( IShellLibrary, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IShellLibrary methods */
    STDMETHOD( LoadLibraryFromItem )( THIS_ IShellItem *, DWORD ) PURE;
    STDMETHOD( LoadLibraryFromKnownFolder )( THIS_ REFKNOWNFOLDERID, DWORD ) PURE;
    STDMETHOD( AddFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( RemoveFolder )( THIS_ IShellItem * ) PURE;
    STDMETHOD( GetFolders )( THIS_ LIBRARYFOLDERFILTER, REFIID, void ** ) PURE;
    STDMETHOD( ResolveFolder )( THIS_ IShellItem *, DWORD, REFIID, void ** ) PURE;
    STDMETHOD( GetDefaultSaveFolder )( THIS_ DEFAULTSAVEFOLDERTYPE, REFIID, void ** ) PURE;
    STDMETHOD( SetDefaultSaveFolder )( THIS_ DEFAULTSAVEFOLDERTYPE, IShellItem * ) PURE;
    STDMETHOD( GetOptions )( THIS_ LIBRARYOPTIONFLAGS * ) PURE;
    STDMETHOD( SetOptions )( THIS_ LIBRARYOPTIONFLAGS, LIBRARYOPTIONFLAGS ) PURE;
    STDMETHOD( GetFolderType )( THIS_ FOLDERTYPEID * ) PURE;
    STDMETHOD( SetFolderType )( THIS_ REFFOLDERTYPEID  ) PURE;
    STDMETHOD( GetIcon )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( SetIcon )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Commit )( THIS ) PURE;
    STDMETHOD( Save )( THIS_ IShellItem *, LPCWSTR, LIBRARYSAVEFLAGS, IShellItem ** ) PURE;
    STDMETHOD( SaveInKnownFolder )( THIS_ REFKNOWNFOLDERID, LPCWSTR, LIBRARYSAVEFLAGS, IShellItem ** ) PURE;
};

/* IAssocHandlerInvoker interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IAssocHandlerInvoker
DECLARE_INTERFACE_( IAssocHandlerInvoker, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IAssocHandlerInvoker methods */
    STDMETHOD( SupportsSelection )( THIS ) PURE;
    STDMETHOD( Invoke )( THIS ) PURE;
};
#endif

/* IAssocHandler interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IAssocHandler
DECLARE_INTERFACE_( IAssocHandler, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IAssocHandler methods */
    STDMETHOD( GetName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( GetUIName )( THIS_ LPWSTR * ) PURE;
    STDMETHOD( GetIconLocation )( THIS_ LPWSTR *, int * ) PURE;
    STDMETHOD( IsRecommended )( THIS ) PURE;
    STDMETHOD( MakeDefault )( THIS_ LPCWSTR ) PURE;
    STDMETHOD( Invoke )( THIS_ IDataObject * ) PURE;
    STDMETHOD( CreateInvoker )( THIS_ IDataObject *, IAssocHandlerInvoker ** ) PURE;
};
#endif

/* IEnumAssocHandlers interface */
#if (NTDDI_VERSION >= 0x06000000)
#undef INTERFACE
#define INTERFACE   IEnumAssocHandlers
DECLARE_INTERFACE_( IEnumAssocHandlers, IUnknown ) {
    /* IUnknown methods */
    STDMETHOD( QueryInterface )( THIS_ REFIID, void ** ) PURE;
    STDMETHOD_( ULONG, AddRef )( THIS ) PURE;
    STDMETHOD_( ULONG, Release )( THIS ) PURE;

    /* IEnumAssocHandlers methods */
    STDMETHOD( Next )( THIS_ ULONG, IAssocHandler **, ULONG * ) PURE;
};
#endif

/* C object macros */
#if (!defined( __cplusplus ) || defined( CINTERFACE )) && defined( COBJMACROS )
    #define IContextMenu_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IContextMenu_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IContextMenu_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IContextMenu_QueryContextMenu( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->QueryContextMenu( x, p1, p2, p3, p4, p5 )
    #define IContextMenu_InvokeCommand( x, p ) \
        (x)->lpVtbl->InvokeCommand( x, p )
    #define IContextMenu_GetCommandString( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetCommandString( x, p1, p2, p3, p4, p5 )
    #define IContextMenu2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IContextMenu2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IContextMenu2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IContextMenu2_QueryContextMenu( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->QueryContextMenu( x, p1, p2, p3, p4, p5 )
    #define IContextMenu2_InvokeCommand( x, p ) \
        (x)->lpVtbl->InvokeCommand( x, p )
    #define IContextMenu2_GetCommandString( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetCommandString( x, p1, p2, p3, p4, p5 )
    #define IContextMenu2_HandleMenuMsg( x, p1, p2, p3 ) \
        (x)->lpVtbl->HandleMenuMsg( x, p1, p2, p3 )
    #define IContextMenu3_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IContextMenu3_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IContextMenu3_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IContextMenu3_QueryContextMenu( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->QueryContextMenu( x, p1, p2, p3, p4, p5 )
    #define IContextMenu3_InvokeCommand( x, p ) \
        (x)->lpVtbl->InvokeCommand( x, p )
    #define IContextMenu3_GetCommandString( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetCommandString( x, p1, p2, p3, p4, p5 )
    #define IContextMenu3_HandleMenuMsg( x, p1, p2, p3 ) \
        (x)->lpVtbl->HandleMenuMsg( x, p1, p2, p3 )
    #define IContextMenu3_HandleMenuMsg2( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->HandleMenuMsg2( x, p1, p2, p3, p4 )
    #define IExecuteCommand_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IExecuteCommand_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IExecuteCommand_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IExecuteCommand_SetKeyState( x, p ) \
        (x)->lpVtbl->SetKeyState( x, p )
    #define IExecuteCommand_SetParameters( x, p ) \
        (x)->lpVtbl->SetParameters( x, p )
    #define IExecuteCommand_SetPosition( x, p ) \
        (x)->lpVtbl->SetPosition( x, p )
    #define IExecuteCommand_SetShowWindow( x, p ) \
        (x)->lpVtbl->SetShowWindow( x, p )
    #define IExecuteCommand_SetNoShowUI( x, p ) \
        (x)->lpVtbl->SetNoShowUI( x, p )
    #define IExecuteCommand_SetDirectory( x, p ) \
        (x)->lpVtbl->SetDirectory( x, p )
    #define IExecuteCommand_Execute( x ) \
        (x)->lpVtbl->Execute( x )
    #define IPersistFolder_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IPersistFolder_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IPersistFolder_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IPersistFolder_GetClassID( x, p ) \
        (x)->lpVtbl->GetClassID( x, p )
    #define IPersistFolder_Initialize( x, p ) \
        (x)->lpVtbl->Initialize( x, p )
    #if (_WIN32_IE >= 0x0400)
        #define IRunnableTask_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IRunnableTask_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IRunnableTask_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IRunnableTask_Run( x ) \
            (x)->lpVtbl->Run( x )
        #define IRunnableTask_Kill( x, p ) \
            (x)->lpVtbl->Kill( x, p )
        #define IRunnableTask_Suspend( x ) \
            (x)->lpVtbl->Suspend( x )
        #define IRunnableTask_Resume( x ) \
            (x)->lpVtbl->Resume( x )
        #define IRunnableTask_IsRunning( x ) \
            (x)->lpVtbl->IsRunning( x )
        #define IShellTaskScheduler_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IShellTaskScheduler_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IShellTaskScheduler_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IShellTaskScheduler_AddTask( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->AddTask( x, p1, p2, p3, p4 )
        #define IShellTaskScheduler_RemoveTasks( x, p1, p2, p3 ) \
            (x)->lpVtbl->RemoveTasks( x, p1, p2, p3 )
        #define IShellTaskScheduler_CountTasks( x, p ) \
            (x)->lpVtbl->CountTasks( x, p )
        #define IShellTaskScheduler_Status( x, p1, p2 ) \
            (x)->lpVtbl->Status( x, p1, p2 )
        #define IQueryCodePage_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IQueryCodePage_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IQueryCodePage_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IQueryCodePage_GetCodePage( x, p ) \
            (x)->lpVtbl->GetCodePage( x, p )
        #define IQueryCodePage_SetCodePage( x, p ) \
            (x)->lpVtbl->SetCodePage( x, p )
        #define IPersistFolder2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IPersistFolder2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IPersistFolder2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IPersistFolder2_GetClassID( x, p ) \
            (x)->lpVtbl->GetClassID( x, p )
        #define IPersistFolder2_Initialize( x, p ) \
            (x)->lpVtbl->Initialize( x, p )
        #define IPersistFolder2_GetCurFolder( x, p ) \
            (x)->lpVtbl->GetCurFolder( x, p )
    #endif
    #if (_WIN32_IE >= 0x0500)
        #define IPersistFolder3_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IPersistFolder3_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IPersistFolder3_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IPersistFolder3_GetClassID( x, p ) \
            (x)->lpVtbl->GetClassID( x, p )
        #define IPersistFolder3_Initialize( x, p ) \
            (x)->lpVtbl->Initialize( x, p )
        #define IPersistFolder3_GetCurFolder( x, p ) \
            (x)->lpVtbl->GetCurFolder( x, p )
        #define IPersistFolder3_InitializeEx( x, p1, p2, p3 ) \
            (x)->lpVtbl->InitializeEx( x, p1, p2, p3 )
        #define IPersistFolder3_GetFolderTargetInfo( x, p ) \
            (x)->lpVtbl->GetFolderTargetInfo( x, p )
    #endif
    #if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0600)
        #define IPersistIDList_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IPersistIDList_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IPersistIDList_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IPersistIDList_GetClassID( x, p ) \
            (x)->lpVtbl->GetClassID( x, p )
        #define IPersistIDList_SetIDList( x, p ) \
            (x)->lpVtbl->SetIDList( x, p )
        #define IPersistIDList_GetIDList( x, p ) \
            (x)->lpVtbl->GetIDList( x, p )
    #endif
    #define IEnumIDList_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumIDList_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumIDList_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumIDList_Next( x, p1, p2, p3 ) \
        (x)->lpVtbl->Next( x, p1, p2, p3 )
    #define IEnumIDList_Skip( x, p ) \
        (x)->lpVtbl->Skip( x, p )
    #define IEnumIDList_Reset( x ) \
        (x)->lpVtbl->Reset( x )
    #define IEnumIDList_Clone( x, p ) \
        (x)->lpVtbl->Clone( x, p )
    #define IEnumIDFullList_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumIDFullList_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumIDFullList_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumIDFullList_Next( x, p1, p2, p3 ) \
        (x)->lpVtbl->Next( x, p1, p2, p3 )
    #define IEnumIDFullList_Skip( x, p ) \
        (x)->lpVtbl->Skip( x, p )
    #define IEnumIDFullList_Reset( x ) \
        (x)->lpVtbl->Reset( x )
    #define IEnumIDFullList_Clone( x, p ) \
        (x)->lpVtbl->Clone( x, p )
    #if (NTDDI_VERSION >= 0x06010000)
        #define IObjectWithFolderEnumMode_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IObjectWithFolderEnumMode_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IObjectWithFolderEnumMode_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IObjectWithFolderEnumMode_SetMode( x, p ) \
            (x)->lpVtbl->SetMode( x, p )
        #define IObjectWithFolderEnumMode_GetMode( x, p ) \
            (x)->lpVtbl->GetMode( x, p )
        #define IParseAndCreateItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IParseAndCreateItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IParseAndCreateItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IParseAndCreateItem_SetItem( x, p ) \
            (x)->lpVtbl->SetItem( x, p )
        #define IParseAndCreateItem_GetItem( x, p1, p2 ) \
            (x)->lpVtbl->GetItem( x, p1, p2 )
    #endif
    #define IShellFolder_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellFolder_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellFolder_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellFolder_ParseDisplayName( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->ParseDisplayName( x, p1, p2, p3, p4, p5, p6 )
    #define IShellFolder_EnumObjects( x, p1, p2, p3 ) \
        (x)->lpVtbl->EnumObjects( x, p1, p2, p3 )
    #define IShellFolder_BindToObject( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToObject( x, p1, p2, p3, p4 )
    #define IShellFolder_BindToStorage( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToStorage( x, p1, p2, p3, p4 )
    #define IShellFolder_CompareIDs( x, p1, p2, p3 ) \
        (x)->lpVtbl->CompareIDs( x, p1, p2, p3 )
    #define IShellFolder_CreateViewObject( x, p1, p2, p3 ) \
        (x)->lpVtbl->CreateViewObject( x, p1, p2, p3 )
    #define IShellFolder_GetAttributesOf( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetAttributesOf( x, p1, p2, p3 )
    #define IShellFolder_GetUIObjectOf( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->GetUIObjectOf( x, p1, p2, p3, p4, p5, p6 )
    #define IShellFolder_GetDisplayNameOf( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDisplayNameOf( x, p1, p2, p3 )
    #define IShellFolder_SetNameOf( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->SetNameOf( x, p1, p2, p3, p4, p5 )
    #define IEnumExtraSearch_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumExtraSearch_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumExtraSearch_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumExtraSearch_Next( x, p1, p2, p3 ) \
        (x)->lpVtbl->Next( x, p1, p2, p3 )
    #define IEnumExtraSearch_Skip( x, p ) \
        (x)->lpVtbl->Skip( x, p )
    #define IEnumExtraSearch_Reset( x ) \
        (x)->lpVtbl->Reset( x )
    #define IEnumExtraSearch_Clone( x, p ) \
        (x)->lpVtbl->Clone( x, p )
    #define IShellFolder2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellFolder2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellFolder2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellFolder2_ParseDisplayName( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->ParseDisplayName( x, p1, p2, p3, p4, p5, p6 )
    #define IShellFolder2_EnumObjects( x, p1, p2, p3 ) \
        (x)->lpVtbl->EnumObjects( x, p1, p2, p3 )
    #define IShellFolder2_BindToObject( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToObject( x, p1, p2, p3, p4 )
    #define IShellFolder2_BindToStorage( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToStorage( x, p1, p2, p3, p4 )
    #define IShellFolder2_CompareIDs( x, p1, p2, p3 ) \
        (x)->lpVtbl->CompareIDs( x, p1, p2, p3 )
    #define IShellFolder2_CreateViewObject( x, p1, p2, p3 ) \
        (x)->lpVtbl->CreateViewObject( x, p1, p2, p3 )
    #define IShellFolder2_GetAttributesOf( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetAttributesOf( x, p1, p2, p3 )
    #define IShellFolder2_GetUIObjectOf( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->GetUIObjectOf( x, p1, p2, p3, p4, p5, p6 )
    #define IShellFolder2_GetDisplayNameOf( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDisplayNameOf( x, p1, p2, p3 )
    #define IShellFolder2_SetNameOf( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->SetNameOf( x, p1, p2, p3, p4, p5 )
    #define IShellFolder2_GetDefaultSearchGUID( x, p ) \
        (x)->lpVtbl->GetDefaultSearchGUID( x, p )
    #define IShellFolder2_EnumSearches( x, p ) \
        (x)->lpVtbl->EnumSearches( x, p )
    #define IShellFolder2_GetDefaultColumn( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDefaultColumn( x, p1, p2, p3 )
    #define IShellFolder2_GetDefaultColumnState( x, p1, p2 ) \
        (x)->lpVtbl->GetDefaultColumnState( x, p1, p2 )
    #define IShellFolder2_GetDetailsEx( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDetailsEx( x, p1, p2, p3 )
    #define IShellFolder2_GetDetailsOf( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDetailsOf( x, p1, p2, p3 )
    #define IShellFolder2_MapColumnToSCID( x, p1, p2 ) \
        (x)->lpVtbl->MapColumnToSCID( x, p1, p2 )
    #define IFolderViewOptions_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IFolderViewOptions_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IFolderViewOptions_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IFolderViewOptions_SetFolderViewOptions( x, p1, p2 ) \
        (x)->lpVtbl->SetFolderViewOptions( x, p1, p2 )
    #define IFolderViewOptions_GetFolderViewOptions( x, p ) \
        (x)->lpVtbl->GetFolderViewOptions( x, p )
    #define IShellView_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellView_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellView_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellView_GetWindow( x, p ) \
        (x)->lpVtbl->GetWindow( x, p )
    #define IShellView_ContextSensitiveHelp( x, p ) \
        (x)->lpVtbl->ContextSensitiveHelp( x, p )
    #define IShellView_TranslateAccelerator( x, p ) \
        (x)->lpVtbl->TranslateAccelerator( x, p )
    #define IShellView_EnableModeless( x, p ) \
        (x)->lpVtbl->EnableModeless( x, p )
    #define IShellView_UIActivate( x, p ) \
        (x)->lpVtbl->UIActivate( x, p )
    #define IShellView_Refresh( x ) \
        (x)->lpVtbl->Refresh( x )
    #define IShellView_CreateViewObject( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->CreateViewObject( x, p1, p2, p3, p4, p5 )
    #define IShellView_DestroyViewObject( x ) \
        (x)->lpVtbl->DestroyViewObject( x )
    #define IShellView_GetCurrentInfo( x, p ) \
        (x)->lpVtbl->GetCurrentInfo( x, p )
    #define IShellView_AddPropertySheetPages( x, p1, p2, p3 ) \
        (x)->lpVtbl->AddPropertySheetPages( x, p1, p2, p3 )
    #define IShellView_SaveViewState( x ) \
        (x)->lpVtbl->SaveViewState( x )
    #define IShellView_SelectItem( x, p1, p2 ) \
        (x)->lpVtbl->SelectItem( x, p1, p2 )
    #define IShellView_GetItemObject( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetItemObject( x, p1, p2, p3 )
    #define IShellView2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellView2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellView2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellView2_GetWindow( x, p ) \
        (x)->lpVtbl->GetWindow( x, p )
    #define IShellView2_ContextSensitiveHelp( x, p ) \
        (x)->lpVtbl->ContextSensitiveHelp( x, p )
    #define IShellView2_TranslateAccelerator( x, p ) \
        (x)->lpVtbl->TranslateAccelerator( x, p )
    #define IShellView2_EnableModeless( x, p ) \
        (x)->lpVtbl->EnableModeless( x, p )
    #define IShellView2_UIActivate( x, p ) \
        (x)->lpVtbl->UIActivate( x, p )
    #define IShellView2_Refresh( x ) \
        (x)->lpVtbl->Refresh( x )
    #define IShellView2_CreateViewObject( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->CreateViewObject( x, p1, p2, p3, p4, p5 )
    #define IShellView2_DestroyViewObject( x ) \
        (x)->lpVtbl->DestroyViewObject( x )
    #define IShellView2_GetCurrentInfo( x, p ) \
        (x)->lpVtbl->GetCurrentInfo( x, p )
    #define IShellView2_AddPropertySheetPages( x, p1, p2, p3 ) \
        (x)->lpVtbl->AddPropertySheetPages( x, p1, p2, p3 )
    #define IShellView2_SaveViewState( x ) \
        (x)->lpVtbl->SaveViewState( x )
    #define IShellView2_SelectItem( x, p1, p2 ) \
        (x)->lpVtbl->SelectItem( x, p1, p2 )
    #define IShellView2_GetItemObject( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetItemObject( x, p1, p2, p3 )
    #define IShellView2_GetView( x, p1, p2 ) \
        (x)->lpVtbl->GetView( x, p1, p2 )
    #define IShellView2_CreateViewWindow2( x, p ) \
        (x)->lpVtbl->CreateViewWindow2( x, p )
    #define IShellView2_HandleResume( x, p ) \
        (x)->lpVtbl->HandleResume( x, p )
    #define IShellView2_SelectAndPositionItem( x, p1, p2, p3 ) \
        (x)->lpVtbl->SelectAndPositionItem( x, p1, p2, p3 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define IShellView3_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IShellView3_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IShellView3_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IShellView3_GetWindow( x, p ) \
            (x)->lpVtbl->GetWindow( x, p )
        #define IShellView3_ContextSensitiveHelp( x, p ) \
            (x)->lpVtbl->ContextSensitiveHelp( x, p )
        #define IShellView3_TranslateAccelerator( x, p ) \
            (x)->lpVtbl->TranslateAccelerator( x, p )
        #define IShellView3_EnableModeless( x, p ) \
            (x)->lpVtbl->EnableModeless( x, p )
        #define IShellView3_UIActivate( x, p ) \
            (x)->lpVtbl->UIActivate( x, p )
        #define IShellView3_Refresh( x ) \
            (x)->lpVtbl->Refresh( x )
        #define IShellView3_CreateViewObject( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->CreateViewObject( x, p1, p2, p3, p4, p5 )
        #define IShellView3_DestroyViewObject( x ) \
            (x)->lpVtbl->DestroyViewObject( x )
        #define IShellView3_GetCurrentInfo( x, p ) \
            (x)->lpVtbl->GetCurrentInfo( x, p )
        #define IShellView3_AddPropertySheetPages( x, p1, p2, p3 ) \
            (x)->lpVtbl->AddPropertySheetPages( x, p1, p2, p3 )
        #define IShellView3_SaveViewState( x ) \
            (x)->lpVtbl->SaveViewState( x )
        #define IShellView3_SelectItem( x, p1, p2 ) \
            (x)->lpVtbl->SelectItem( x, p1, p2 )
        #define IShellView3_GetItemObject( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetItemObject( x, p1, p2, p3 )
        #define IShellView3_GetView( x, p1, p2 ) \
            (x)->lpVtbl->GetView( x, p1, p2 )
        #define IShellView3_CreateViewWindow2( x, p ) \
            (x)->lpVtbl->CreateViewWindow2( x, p )
        #define IShellView3_HandleResume( x, p ) \
            (x)->lpVtbl->HandleResume( x, p )
        #define IShellView3_SelectAndPositionItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->SelectAndPositionItem( x, p1, p2, p3 )
        #define IShellView2_CreateViewWindow3( x, p1, p2, p3, p4, p5, p6, p7, p8, p9 ) \
            (x)->lpVtbl->CreateViewWindow2( x, p1, p2, p3, p4, p5, p6, p7, p8, p9 )
    #endif
    #define IFolderView_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IFolderView_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IFolderView_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IFolderView_GetCurrentViewMode( x, p ) \
        (x)->lpVtbl->GetCurrentViewMode( x, p )
    #define IFolderView_SetCurrentViewMode( x, p ) \
        (x)->lpVtbl->SetCurrentViewMode( x, p )
    #define IFolderView_GetFolder( x, p1, p2 ) \
        (x)->lpVtbl->GetFolder( x, p1, p2 )
    #define IFolderView_Item( x, p1, p2 ) \
        (x)->lpVtbl->Item( x, p1, p2 )
    #define IFolderView_ItemCount( x, p1, p2 ) \
        (x)->lpVtbl->ItemCount( x, p1, p2 )
    #define IFolderView_Items( x, p1, p2, p3 ) \
        (x)->lpVtbl->Items( x, p1, p2, p3 )
    #define IFolderView_GetSelectionMarkedItem( x, p ) \
        (x)->lpVtbl->GetSelectionMarkedItem( x, p )
    #define IFolderView_GetFocusedItem( x, p ) \
        (x)->lpVtbl->GetFocusedItem( x, p )
    #define IFolderView_GetItemPosition( x, p1, p2 ) \
        (x)->lpVtbl->GetItemPosition( x, p1, p2 )
    #define IFolderView_GetSpacing( x, p ) \
        (x)->lpVtbl->GetSpacing( x, p )
    #define IFolderView_GetDefaultSpacing( x, p ) \
        (x)->lpVtbl->GetDefaultSpacing( x, p )
    #define IFolderView_GetAutoArrange( x ) \
        (x)->lpVtbl->GetAutoArrange( x )
    #define IFolderView_SelectItem( x, p1, p2 ) \
        (x)->lpVtbl->SelectItem( x, p1, p2 )
    #define IFolderView_SelectAndPositionItems( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->SelectAndPositionItems( x, p1, p2, p3, p4 )
    #if (NTDDI_VERSION >= 0x06010000)
        #define ISearchBoxInfo_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ISearchBoxInfo_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ISearchBoxInfo_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ISearchBoxInfo_GetCondition( x, p1, p2 ) \
            (x)->lpVtbl->GetCondition( x, p1, p2 )
        #define ISearchBoxInfo_GetText( x, p ) \
            (x)->lpVtbl->GetText( x, p )
    #endif
    #if (NTDDI_VERSION >= 0x06000000) || (_WIN32_IE >= 0x0700)
        #define IFolderView2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFolderView2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFolderView2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFolderView2_GetCurrentViewMode( x, p ) \
            (x)->lpVtbl->GetCurrentViewMode( x, p )
        #define IFolderView2_SetCurrentViewMode( x, p ) \
            (x)->lpVtbl->SetCurrentViewMode( x, p )
        #define IFolderView2_GetFolder( x, p1, p2 ) \
            (x)->lpVtbl->GetFolder( x, p1, p2 )
        #define IFolderView2_Item( x, p1, p2 ) \
            (x)->lpVtbl->Item( x, p1, p2 )
        #define IFolderView2_ItemCount( x, p1, p2 ) \
            (x)->lpVtbl->ItemCount( x, p1, p2 )
        #define IFolderView2_Items( x, p1, p2, p3 ) \
            (x)->lpVtbl->Items( x, p1, p2, p3 )
        #define IFolderView2_GetSelectionMarkedItem( x, p ) \
            (x)->lpVtbl->GetSelectionMarkedItem( x, p )
        #define IFolderView2_GetFocusedItem( x, p ) \
            (x)->lpVtbl->GetFocusedItem( x, p )
        #define IFolderView2_GetItemPosition( x, p1, p2 ) \
            (x)->lpVtbl->GetItemPosition( x, p1, p2 )
        #define IFolderView2_GetSpacing( x, p ) \
            (x)->lpVtbl->GetSpacing( x, p )
        #define IFolderView2_GetDefaultSpacing( x, p ) \
            (x)->lpVtbl->GetDefaultSpacing( x, p )
        #define IFolderView2_GetAutoArrange( x ) \
            (x)->lpVtbl->GetAutoArrange( x )
        #define IFolderView2_SelectItem( x, p1, p2 ) \
            (x)->lpVtbl->SelectItem( x, p1, p2 )
        #define IFolderView2_SelectAndPositionItems( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->SelectAndPositionItems( x, p1, p2, p3, p4 )
        #define IFolderView2_SetGroupBy( x, p1, p2 ) \
            (x)->lpVtbl->SetGroupBy( x, p1, p2 )
        #define IFolderView2_GetGroupBy( x, p1, p2 ) \
            (x)->lpVtbl->GetGroupBy( x, p1, p2 )
        #define IFolderView2_SetViewProperty( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetViewProperty( x, p1, p2, p3 )
        #define IFolderView2_GetViewProperty( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetViewProperty( x, p1, p2, p3 )
        #define IFolderView2_SetTileViewProperties( x, p1, p2 ) \
            (x)->lpVtbl->SetTileViewProperties( x, p1, p2 )
        #define IFolderView2_SetExtendedTileViewProperties( x, p1, p2 ) \
            (x)->lpVtbl->SetExtendedTileViewProperties( x, p1, p2 )
        #define IFolderView2_SetText( x, p1, p2 ) \
            (x)->lpVtbl->SetText( x, p1, p2 )
        #define IFolderView2_SetCurrentFolderFlags( x, p1, p2 ) \
            (x)->lpVtbl->SetCurrentFolderFlags( x, p1, p2 )
        #define IFolderView2_GetCurrentFolderFlags( x, p ) \
            (x)->lpVtbl->GetCurrentFolderFlags( x, p )
        #define IFolderView2_GetSortColumnCount( x, p ) \
            (x)->lpVtbl->GetSortColumnCount( x, p )
        #define IFolderView2_SetSortColumns( x, p1, p2 ) \
            (x)->lpVtbl->SetSortColumns( x, p1, p2 )
        #define IFolderView2_GetSortColumns( x, p1, p2 ) \
            (x)->lpVtbl->GetSortColumns( x, p1, p2 )
        #define IFolderView2_GetItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetItem( x, p1, p2, p3 )
        #define IFolderView2_GetVisibleItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetVisibleItem( x, p1, p2, p3 )
        #define IFolderView2_GetSelectedItem( x, p1, p2 ) \
            (x)->lpVtbl->GetSelectedItem( x, p1, p2 )
        #define IFolderView2_GetSelection( x, p1, p2 ) \
            (x)->lpVtbl->GetSelection( x, p1, p2 )
        #define IFolderView2_GetSelectionState( x, p1, p2 ) \
            (x)->lpVtbl->GetSelectionState( x, p1, p2 )
        #define IFolderView2_InvokeVerbOnSelection( x, p ) \
            (x)->lpVtbl->InvokeVerbOnSelection( x, p )
        #define IFolderView2_SetViewModeAndIconSize( x, p1, p2 ) \
            (x)->lpVtbl->SetViewModeAndIconSize( x, p1, p2 )
        #define IFolderView2_GetViewModeAndIconSize( x, p1, p2 ) \
            (x)->lpVtbl->GetViewModeAndIconSize( x, p1, p2 )
        #define IFolderView2_SetGroupSubsetCount( x, p ) \
            (x)->lpVtbl->SetGroupSubsetCount( x, p )
        #define IFolderView2_GetGroupSubsetCount( x, p ) \
            (x)->lpVtbl->GetGroupSubsetCount( x, p )
        #define IFolderView2_SetRedraw( x, p ) \
            (x)->lpVtbl->SetRedraw( x, p )
        #define IFolderView2_IsMoveInSameFolder( x ) \
            (x)->lpVtbl->IsMoveInSameFolder( x )
        #define IFolderView2_DoRename( x ) \
            (x)->lpVtbl->DoRename( x )
    #endif
    #if (NTDDI_VERSION >= 0x06000000)
        #define IFolderViewSettings_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFolderViewSettings_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFolderViewSettings_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFolderViewSettings_GetColumnPropertyList( x, p1, p2 ) \
            (x)->lpVtbl->GetColumnPropertyList( x, p1, p2 )
        #define IFolderViewSettings_GetGroupByProperty( x, p1, p2 ) \
            (x)->lpVtbl->GetGroupByProperty( x, p1, p2 )
        #define IFolderViewSettings_GetViewMode( x, p ) \
            (x)->lpVtbl->GetViewMode( x, p )
        #define IFolderViewSettings_GetIconSize( x, p ) \
            (x)->lpVtbl->GetIconSize( x, p )
        #define IFolderViewSettings_GetFolderFlags( x, p1, p2 ) \
            (x)->lpVtbl->GetFolderFlags( x, p1, p2 )
        #define IFolderViewSettings_GetSortColumns( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetSortColumns( x, p1, p2, p3 )
        #define IFolderViewSettings_GetGroupSubsetCount( x, p ) \
            (x)->lpVtbl->GetGroupSubsetCount( x, p )
    #endif
    #if (_WIN32_IE >= 0x0700)
        #define IPreviewHandlerVisuals_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IPreviewHandlerVisuals_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IPreviewHandlerVisuals_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IPreviewHandlerVisuals_SetBackgroundColor( x, p ) \
            (x)->lpVtbl->SetBackgroundColor( x, p )
        #define IPreviewHandlerVisuals_SetFont( x, p ) \
            (x)->lpVtbl->SetFont( x, p )
        #define IPreviewHandlerVisuals_SetTextColor( x, p ) \
            (x)->lpVtbl->SetTextColor( x, p )
        #define IVisualProperties_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IVisualProperties_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IVisualProperties_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IVisualProperties_SetWatermark( x, p1, p2 ) \
            (x)->lpVtbl->SetWatermark( x, p1, p2 )
        #define IVisualProperties_SetColor( x, p1, p2 ) \
            (x)->lpVtbl->SetColor( x, p1, p2 )
        #define IVisualProperties_GetColor( x, p1, p2 ) \
            (x)->lpVtbl->GetColor( x, p1, p2 )
        #define IVisualProperties_SetItemHeight( x, p ) \
            (x)->lpVtbl->SetItemHeight( x, p )
        #define IVisualProperties_GetItemHeight( x, p ) \
            (x)->lpVtbl->GetItemHeight( x, p )
        #define IVisualProperties_SetFont( x, p1, p2 ) \
            (x)->lpVtbl->SetFont( x, p1, p2 )
        #define IVisualProperties_GetFont( x, p ) \
            (x)->lpVtbl->GetFont( x, p )
        #define IVisualProperties_SetTheme( x, p1, p2 ) \
            (x)->lpVtbl->SetTheme( x, p1, p2 )
    #endif
    #define ICommDlgBrowser_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ICommDlgBrowser_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ICommDlgBrowser_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ICommDlgBrowser_OnDefaultCommand( x, p ) \
        (x)->lpVtbl->OnDefaultCommand( x, p )
    #define ICommDlgBrowser_OnStateChange( x, p1, p2 ) \
        (x)->lpVtbl->OnStateChange( x, p1, p2 )
    #define ICommDlgBrowser_IncludeObject( x, p1, p2 ) \
        (x)->lpVtbl->IncludeObject( x, p1, p2 )
    #if (NTDDI_VERSION >= 0x05000000)
        #define ICommDlgBrowser2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICommDlgBrowser2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICommDlgBrowser2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICommDlgBrowser2_OnDefaultCommand( x, p ) \
            (x)->lpVtbl->OnDefaultCommand( x, p )
        #define ICommDlgBrowser2_OnStateChange( x, p1, p2 ) \
            (x)->lpVtbl->OnStateChange( x, p1, p2 )
        #define ICommDlgBrowser2_IncludeObject( x, p1, p2 ) \
            (x)->lpVtbl->IncludeObject( x, p1, p2 )
        #define ICommDlgBrowser2_Notify( x, p1, p2 ) \
            (x)->lpVtbl->Notify( x, p1, p2 )
        #define ICommDlgBrowser2_GetDefaultMenuText( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetDefaultMenuText( x, p1, p2, p3 )
        #define ICommDlgBrowser2_GetViewFlags( x, p ) \
            (x)->lpVtbl->GetViewFlags( x, p )
    #endif
    #if (_WIN32_IE >= 0x0700)
        #define ICommDlgBrowser3_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICommDlgBrowser3_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICommDlgBrowser3_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICommDlgBrowser3_OnDefaultCommand( x, p ) \
            (x)->lpVtbl->OnDefaultCommand( x, p )
        #define ICommDlgBrowser3_OnStateChange( x, p1, p2 ) \
            (x)->lpVtbl->OnStateChange( x, p1, p2 )
        #define ICommDlgBrowser3_IncludeObject( x, p1, p2 ) \
            (x)->lpVtbl->IncludeObject( x, p1, p2 )
        #define ICommDlgBrowser3_Notify( x, p1, p2 ) \
            (x)->lpVtbl->Notify( x, p1, p2 )
        #define ICommDlgBrowser3_GetDefaultMenuText( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetDefaultMenuText( x, p1, p2, p3 )
        #define ICommDlgBrowser3_GetViewFlags( x, p ) \
            (x)->lpVtbl->GetViewFlags( x, p )
        #define ICommDlgBrowser3_OnColumnClicked( x, p1, p2 ) \
            (x)->lpVtbl->OnColumnClicked( x, p1, p2 )
        #define ICommDlgBrowser3_GetCurrentFilter( x, p1, p2 ) \
            (x)->lpVtbl->GetCurrentFilter( x, p1, p2 )
        #define ICommDlgBrowser3_OnPreViewCreated( x, p ) \
            (x)->lpVtbl->GetPreViewCreated( x, p )
        #define IColumnManager_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IColumnManager_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IColumnManager_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IColumnManager_SetColumnInfo( x, p1, p2 ) \
            (x)->lpVtbl->SetColumnInfo( x, p1, p2 )
        #define IColumnManager_GetColumnInfo( x, p1, p2 ) \
            (x)->lpVtbl->GetColumnInfo( x, p1, p2 )
        #define IColumnManager_GetColumnCount( x, p1, p2 ) \
            (x)->lpVtbl->GetColumnCount( x, p1, p2 )
        #define IColumnManager_GetColumns( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetColumns( x, p1, p2, p3 )
        #define IColumnManager_SetColumns( x, p1, p2 ) \
            (x)->lpVtbl->SetColumns( x, p1, p2 )
    #endif
    #define IFolderFilterSite_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IFolderFilterSite_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IFolderFilterSite_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IFolderFilterSite_SetFilter( x, p ) \
        (x)->lpVtbl->SetFilter( x, p )
    #define IFolderFilter_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IFolderFilter_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IFolderFilter_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IFolderFilter_ShouldShow( x, p1, p2, p3 ) \
        (x)->lpVtbl->ShouldShow( x, p1, p2, p3 )
    #define IFolderFilter_GetEnumFlags( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetEnumFlags( x, p1, p2, p3, p4 )
    #define IInputObjectSite_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInputObjectSite_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInputObjectSite_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInputObjectSite_OnFocusChangeIS( x, p1, p2 ) \
        (x)->lpVtbl->OnFocusChangeIS( x, p1, p2 )
    #define IInputObject_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInputObject_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInputObject_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInputObject_UIActivateIO( x, p1, p2 ) \
        (x)->lpVtbl->UIActivateIO( x, p1, p2 )
    #define IInputObject_HasFocusIO( x ) \
        (x)->lpVtbl->HasFocusIO( x )
    #define IInputObject_TranslateAcceleratorIO( x, p ) \
        (x)->lpVtbl->TranslateAcceleratorIO( x, p )
    #define IInputObject2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInputObject2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInputObject2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInputObject2_UIActivateIO( x, p1, p2 ) \
        (x)->lpVtbl->UIActivateIO( x, p1, p2 )
    #define IInputObject2_HasFocusIO( x ) \
        (x)->lpVtbl->HasFocusIO( x )
    #define IInputObject2_TranslateAcceleratorIO( x, p ) \
        (x)->lpVtbl->TranslateAcceleratorIO( x, p )
    #define IInputObject2_TranslateAcceleratorGlobal( x, p ) \
        (x)->lpVtbl->TranslateAcceleratorGlobal( x, p )
    #define IShellIcon_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellIcon_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellIcon_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellIcon_GetIconOf( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetIconOf( x, p1, p2, p3 )
    #define IShellBrowser_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellBrowser_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellBrowser_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellBrowser_GetWindow( x, p ) \
        (x)->lpVtbl->GetWindow( x, p )
    #define IShellBrowser_ContextSensitiveHelp( x, p ) \
        (x)->lpVtbl->ContextSensitiveHelp( x, p )
    #define IShellBrowser_InsertMenusSB( x, p1, p2 ) \
        (x)->lpVtbl->InsertMenusSB( x, p1, p2 )
    #define IShellBrowser_SetMenuSB( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetMenuSB( x, p1, p2, p3 )
    #define IShellBrowser_RemoveMenusSB( x, p ) \
        (x)->lpVtbl->RemoveMenusSB( x, p )
    #define IShellBrowser_SetStatusTextSB( x, p ) \
        (x)->lpVtbl->SetStatusTextSB( x, p )
    #define IShellBrowser_EnableModelessSB( x, p ) \
        (x)->lpVtbl->EnableModelessSB( x, p )
    #define IShellBrowser_TranslateAcceleratorSB( x, p1, p2 ) \
        (x)->lpVtbl->TranslateAcceleratorSB( x, p1, p2 )
    #define IShellBrowser_BrowseObject( x, p1, p2 ) \
        (x)->lpVtbl->BrowseObject( x, p1, p2 )
    #define IShellBrowser_GetViewStateStream( x, p1, p2 ) \
        (x)->lpVtbl->GetViewStateStream( x, p1, p2 )
    #define IShellBrowser_GetControlWindow( x, p1, p2 ) \
        (x)->lpVtbl->GetControlWindow( x, p1, p2 )
    #define IShellBrowser_SendControlMsg( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->SendControlMsg( x, p1, p2, p3, p4, p5 )
    #define IShellBrowser_QueryActiveShellView( x, p ) \
        (x)->lpVtbl->QueryActiveShellView( x, p )
    #define IShellBrowser_OnViewWindowActive( x, p ) \
        (x)->lpVtbl->OnViewWindowActive( x, p )
    #define IShellBrowser_SetToolbarItems( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetToolbarItems( x, p1, p2, p3 )
    #define IProfferService_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IProfferService_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IProfferService_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IProfferService_ProfferService( x, p1, p2, p3 ) \
        (x)->lpVtbl->ProfferService( x, p1, p2, p3 )
    #define IProfferService_RevokeService( x, p ) \
        (x)->lpVtbl->RevokeService( x, p )
    #define IShellItem_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellItem_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellItem_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellItem_BindToHandler( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToHandler( x, p1, p2, p3, p4 )
    #define IShellItem_GetParent( x, p ) \
        (x)->lpVtbl->GetParent( x, p )
    #define IShellItem_GetDisplayName( x, p1, p2 ) \
        (x)->lpVtbl->GetDisplayName( x, p1, p2 )
    #define IShellItem_GetAttributes( x, p1, p2 ) \
        (x)->lpVtbl->GetAttributes( x, p1, p2 )
    #define IShellItem_Compare( x, p1, p2, p3 ) \
        (x)->lpVtbl->Compare( x, p1, p2, p3 )
    #define IShellItem2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellItem2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellItem2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellItem2_BindToHandler( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToHandler( x, p1, p2, p3, p4 )
    #define IShellItem2_GetParent( x, p ) \
        (x)->lpVtbl->GetParent( x, p )
    #define IShellItem2_GetDisplayName( x, p1, p2 ) \
        (x)->lpVtbl->GetDisplayName( x, p1, p2 )
    #define IShellItem2_GetAttributes( x, p1, p2 ) \
        (x)->lpVtbl->GetAttributes( x, p1, p2 )
    #define IShellItem2_Compare( x, p1, p2, p3 ) \
        (x)->lpVtbl->Compare( x, p1, p2, p3 )
    #define IShellItem2_GetPropertyStore( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetPropertyStore( x, p1, p2, p3 )
    #define IShellItem2_GetPropertyStoreWithCreateObject( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetPropertyStoreWithCreateObject( x, p1, p2, p3, p4 )
    #define IShellItem2_GetPropertyStoreForKeys( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetPropertyStoreForKeys( x, p1, p2, p3, p4, p5 )
    #define IShellItem2_GetPropertyDescriptionList( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetPropertyDescriptionList( x, p1, p2, p3 )
    #define IShellItem2_Update( x, p ) \
        (x)->lpVtbl->Update( x, p )
    #define IShellItem2_GetProperty( x, p1, p2 ) \
        (x)->lpVtbl->GetProperty( x, p1, p2 )
    #define IShellItem2_GetCLSID( x, p1, p2 ) \
        (x)->lpVtbl->GetCLSID( x, p1, p2 )
    #define IShellItem2_GetFileTime( x, p1, p2 ) \
        (x)->lpVtbl->GetFileTime( x, p1, p2 )
    #define IShellItem2_GetInt32( x, p1, p2 ) \
        (x)->lpVtbl->GetInt32( x, p1, p2 )
    #define IShellItem2_GetString( x, p1, p2 ) \
        (x)->lpVtbl->GetString( x, p1, p2 )
    #define IShellItem2_GetUInt32( x, p1, p2 ) \
        (x)->lpVtbl->GetUInt32( x, p1, p2 )
    #define IShellItem2_GetUInt64( x, p1, p2 ) \
        (x)->lpVtbl->GetUInt64( x, p1, p2 )
    #define IShellItem2_GetBool( x, p1, p2 ) \
        (x)->lpVtbl->GetBool( x, p1, p2 )
    #define IShellItemImageFactory_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellItemImageFactory_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellItemImageFactory_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellItemImageFactory_GetImage( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetImage( x, p1, p2, p3 )
    #define IUserAccountChangeCallback_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IUserAccountChangeCallback_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IUserAccountChangeCallback_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IUserAccountChangeCallback_OnPictureChange( x, p ) \
        (x)->lpVtbl->OnPictureChange( x, p )
    #if (NTDDI_VERSION >= 0x05010000)
        #define IEnumShellItems_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IEnumShellItems_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IEnumShellItems_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IEnumShellItems_Next( x, p1, p2, p3 ) \
            (x)->lpVtbl->Next( x, p1, p2, p3 )
        #define IEnumShellItems_Skip( x, p ) \
            (x)->lpVtbl->Skip( x, p )
        #define IEnumShellItems_Reset( x ) \
            (x)->lpVtbl->Reset( x )
        #define IEnumShellItems_Clone( x, p ) \
            (x)->lpVtbl->Clone( x, p )
    #endif
    #if (_WIN32_IE >= 0x0700)
        #define ITransferAdviseSink_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ITransferAdviseSink_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ITransferAdviseSink_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ITransferAdviseSink_UpdateProgress( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->UpdateProgress( x, p1, p2, p3, p4, p5, p6 )
        #define ITransferAdviseSink_UpdateTransferState( x, p ) \
            (x)->lpVtbl->UpdateTransferState( x, p )
        #define ITransferAdviseSink_ConfirmOverwrite( x, p1, p2, p3 ) \
            (x)->lpVtbl->ConfirmOverwrite( x, p1, p2, p3 )
        #define ITransferAdviseSink_ConfirmEncryptionLoss( x, p ) \
            (x)->lpVtbl->ConfirmEncryptionLoss( x, p )
        #define ITransferAdviseSink_FileFailure( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->FailFailure( x, p1, p2, p3, p4, p5 )
        #define ITransferAdviseSink_SubStreamFailure( x, p1, p2, p3 ) \
            (x)->lpVtbl->SubStreamFailure( x, p1, p2, p3 )
        #define ITransferAdviseSink_PropertyFailure( x, p1, p2, p3 ) \
            (x)->lpVtbl->PropertyFailure( x, p1, p2, p3 )
    #endif
    #if (NTDDI_VERSION >= 0x06000000)
        #define ITransferSource_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ITransferSource_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ITransferSource_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ITransferSource_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define ITransferSource_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define ITransferSource_SetProperties( x, p ) \
            (x)->lpVtbl->SetProperties( x, p )
        #define ITransferSource_OpenItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->OpenItem( x, p1, p2, p3, p4 )
        #define ITransferSource_MoveItem( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->MoveItem( x, p1, p2, p3, p4, p5 )
        #define ITransferSource_RecycleItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->RecycleItem( x, p1, p2, p3, p4 )
        #define ITransferSource_RemoveItem( x, p1, p2 ) \
            (x)->lpVtbl->RemoveItem( x, p1, p2 )
        #define ITransferSource_RenameItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->RenameItem( x, p1, p2, p3, p4 )
        #define ITransferSource_LinkItem( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->LinkItem( x, p1, p2, p3, p4, p5 )
        #define ITransferSource_ApplyPropertiesToItem( x, p1, p2 ) \
            (x)->lpVtbl->ApplyPropertiesToItem( x, p1, p2 )
        #define ITransferSource_GetDefaultDestinationName( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetDefaultDestinationName( x, p1, p2, p3 )
        #define ITransferSource_EnterFolder( x, p ) \
            (x)->lpVtbl->EnterFolder( x, p )
        #define ITransferSource_LeaveFolder( x, p ) \
            (x)->lpVtbl->LeaveFolder( x, p )
    #endif
    #define IEnumResources_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumResources_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumResources_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumResources_Next( x, p1, p2, p3 ) \
        (x)->lpVtbl->Next( x, p1, p2, p3 )
    #define IEnumResources_Skip( x, p ) \
        (x)->lpVtbl->Skip( x, p )
    #define IEnumResources_Reset( x ) \
        (x)->lpVtbl->Reset( x )
    #define IEnumResources_Clone( x, p ) \
        (x)->lpVtbl->Clone( x, p )
    #define IShellItemResources_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellItemResources_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellItemResources_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellItemResources_GetAttributes( x, p ) \
        (x)->lpVtbl->GetAttributes( x, p )
    #define IShellItemResources_GetSize( x, p ) \
        (x)->lpVtbl->GetSize( x, p )
    #define IShellItemResources_GetTimes( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetTimes( x, p1, p2, p3 )
    #define IShellItemResources_SetTimes( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetTimes( x, p1, p2, p3 )
    #define IShellItemResources_GetResourceDescription( x, p1, p2 ) \
        (x)->lpVtbl->GetResourceDescription( x, p1, p2 )
    #define IShellItemResources_EnumResources( x, p ) \
        (x)->lpVtbl->EnumResources( x, p )
    #define IShellItemResources_SupportsResource( x, p ) \
        (x)->lpVtbl->SupportsResource( x, p )
    #define IShellItemResources_OpenResource( x, p1, p2, p3 ) \
        (x)->lpVtbl->OpenResource( x, p1, p2, p3 )
    #define IShellItemResources_CreateResource( x, p1, p2, p3 ) \
        (x)->lpVtbl->CreateResource( x, p1, p2, p3 )
    #define IShellItemResources_MarkForDelete( x ) \
        (x)->lpVtbl->MarkForDelete( x )
    #define ITransferDestination_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ITransferDestination_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ITransferDestination_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ITransferDestination_Advise( x, p1, p2 ) \
        (x)->lpVtbl->Advise( x, p1, p2 )
    #define ITransferDestination_Unadvise( x, p ) \
        (x)->lpVtbl->Unadvise( x, p )
    #define ITransferDestination_CreateItem( x, p1, p2, p3, p4, p5, p6, p7, p8 ) \
        (x)->lpVtbl->CreateItem( x, p1, p2, p3, p4, p5, p6, p7, p8 )
    #define IStreamAsync_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IStreamAsync_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IStreamAsync_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IStreamAsync_Read( x, p1, p2, p3 ) \
        (x)->lpVtbl->Read( x, p1, p2, p3 )
    #define IStreamAsync_Write( x, p1, p2, p3 ) \
        (x)->lpVtbl->Write( x, p1, p2, p3 )
    #define IStreamAsync_Seek( x, p1, p2, p3 ) \
        (x)->lpVtbl->Seek( x, p1, p2, p3 )
    #define IStreamAsync_SetSize( x, p ) \
        (x)->lpVtbl->SetSize( x, p )
    #define IStreamAsync_CopyTo( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->CopyTo( x, p1, p2, p3, p4 )
    #define IStreamAsync_Commit( x, p ) \
        (x)->lpVtbl->Commit( x, p )
    #define IStreamAsync_Revert( x ) \
        (x)->lpVtbl->Revert( x )
    #define IStreamAsync_LockRegion( x, p1, p2, p3 ) \
        (x)->lpVtbl->LockRegion( x, p1, p2, p3 )
    #define IStreamAsync_UnlockRegion( x, p1, p2, p3 ) \
        (x)->lpVtbl->UnlockRegion( x, p1, p2, p3 )
    #define IStreamAsync_Stat( x, p1, p2 ) \
        (x)->lpVtbl->Stat( x, p1, p2 )
    #define IStreamAsync_Clone( x, p ) \
        (x)->lpVtbl->Clone( x, p )
    #define IStreamAsync_ReadAsync( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->ReadAsync( x, p1, p2, p3, p4 )
    #define IStreamAsync_WriteAsync( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->WriteAsync( x, p1, p2, p3, p4 )
    #define IStreamAsync_OverlappedResult( x, p1, p2, p3 ) \
        (x)->lpVtbl->OverlappedResult( x, p1, p2, p3 )
    #define IStreamAsync_CancelIo( x ) \
        (x)->lpVtbl->CancelIo( x )
    #define IStreamUnbufferedInfo_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IStreamUnbufferedInfo_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IStreamUnbufferedInfo_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IStreamUnbufferedInfo_GetSectorSize( x, p ) \
        (x)->lpVtbl->GetSectorSize( x, p )
    #if (_WIN32_IE >= 0x0700)
        #define IFileOperationProgressSink_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileOperationProgressSink_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileOperationProgressSink_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileOperationProgressSink_StartOperations( x ) \
            (x)->lpVtbl->StartOperations( x )
        #define IFileOperationProgressSink_FinishOperations( x, p ) \
            (x)->lpVtbl->FinishOperations( x, p )
        #define IFileOperationProgressSink_PreRenameItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->PreRenameItem( x, p1, p2, p3 )
        #define IFileOperationProgressSink_PostRenameItem( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->PostRenameItem( x, p1, p2, p3, p4, p5 )
        #define IFileOperationProgressSink_PreMoveItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->PreMoveItem( x, p1, p2, p3, p4 )
        #define IFileOperationProgressSink_PostMoveItem( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->PostMoveItem( x, p1, p2, p3, p4, p5, p6 )
        #define IFileOperationProgressSink_PreCopyItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->PreCopyItem( x, p1, p2, p3, p4 )
        #define IFileOperationProgressSink_PostCopyItem( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->PostCopyItem( x, p1, p2, p3, p4, p5, p6 )
        #define IFileOperationProgressSink_PreDeleteItem( x, p1, p2 ) \
            (x)->lpVtbl->PreDeleteItem( x, p1, p2 )
        #define IFileOperationProgressSink_PostDeleteItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->PostDeleteItem( x, p1, p2, p3, p4 )
        #define IFileOperationProgressSink_PreNewItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->PreNewItem( x, p1, p2, p3 )
        #define IFileOperationProgressSink_PostNewItem( x, p1, p2, p3, p4, p5, p6, p7 ) \
            (x)->lpVtbl->PostNewItem( x, p1, p2, p3, p4, p5, p6, p7 )
        #define IFileOperationProgressSink_UpdateProgress( x, p1, p2 ) \
            (x)->lpVtbl->UpdateProgress( x, p1, p2 )
        #define IFileOperationProgressSink_ResetTimer( x ) \
            (x)->lpVtbl->ResetTimer( x )
        #define IFileOperationProgressSink_PauseTimer( x ) \
            (x)->lpVtbl->PauseTimer( x )
        #define IFileOperationProgressSink_ResumeTimer( x ) \
            (x)->lpVtbl->ResumeTimer( x )
    #endif
    #define IShellItemArray_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellItemArray_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellItemArray_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellItemArray_BindToHandler( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->BindToHandler( x, p1, p2, p3, p4 )
    #define IShellItemArray_GetPropertyStore( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetPropertyStore( x, p1, p2, p3 )
    #define IShellItemArray_GetPropertyDescriptionList( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetPropertyDescriptionList( x, p1, p2, p3 )
    #define IShellItemArray_GetAttributes( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetAttributes( x, p1, p2, p3 )
    #define IShellItemArray_GetCount( x, p ) \
        (x)->lpVtbl->GetCount( x, p )
    #define IShellItemArray_GetItemAt( x, p1, p2 ) \
        (x)->lpVtbl->GetItemAt( x, p1, p2 )
    #define IShellItemArray_EnumItems( x, p ) \
        (x)->lpVtbl->EnumItems( x, p )
    #define IInitializeWithItem_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInitializeWithItem_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInitializeWithItem_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInitializeWithItem_Initialize( x, p1, p2 ) \
        (x)->lpVtbl->Initialize( x, p1, p2 )
    #define IObjectWithSelection_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IObjectWithSelection_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IObjectWithSelection_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IObjectWithSelection_SetSelection( x, p ) \
        (x)->lpVtbl->SetSelection( x, p )
    #define IObjectWithSelection_GetSelection( x, p1, p2 ) \
        (x)->lpVtbl->GetSelection( x, p1, p2 )
    #define IObjectWithBackReferences_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IObjectWithBackReferences_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IObjectWithBackReferences_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IObjectWithBackReferences_RemoveBackReferences( x ) \
        (x)->lpVtbl->RemoveBackReferences( x )
    #define IPropertyUI_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IPropertyUI_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IPropertyUI_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IPropertyUI_ParsePropertyName( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->ParsePropertyName( x, p1, p2, p3, p4 )
    #define IPropertyUI_GetCanonicalName( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetCanonicalName( x, p1, p2, p3, p4 )
    #define IPropertyUI_GetDisplayName( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetDisplayName( x, p1, p2, p3, p4, p5 )
    #define IPropertyUI_GetPropertyDescription( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetPropertyDescription( x, p1, p2, p3, p4 )
    #define IPropertyUI_GetDefaultWidth( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDefaultWidth( x, p1, p2, p3 )
    #define IPropertyUI_GetFlags( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetFlags( x, p1, p2, p3 )
    #define IPropertyUI_FormatForDisplay( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->FormatForDisplay( x, p1, p2, p3, p4, p5, p6 )
    #define IPropertyUI_GetHelpInfo( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetHelpInfo( x, p1, p2, p3, p4, p5 )
    #if (_WIN32_IE >= 0x0500)
        #define ICategoryProvider_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICategoryProvider_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICategoryProvider_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICategoryProvider_CanCategorizeOnSCID( x, p ) \
            (x)->lpVtbl->CanCategorizeOnSCID( x, p )
        #define ICategoryProvider_GetDefaultCategory( x, p1, p2 ) \
            (x)->lpVtbl->GetDefaultCategory( x, p1, p2 )
        #define ICategoryProvider_GetCategoryForSCID( x, p1, p2 ) \
            (x)->lpVtbl->GetCategoryForSCID( x, p1, p2 )
        #define ICategoryProvider_EnumCategories( x, p ) \
            (x)->lpVtbl->EnumCategories( x, p )
        #define ICategoryProvider_GetCategoryName( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetCategoryName( x, p1, p2, p3 )
        #define ICategoryProvider_CreateCategory( x, p1, p2, p3 ) \
            (x)->lpVtbl->CreateCategory( x, p1, p2, p3 )
        #define ICategorizer_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICategorizer_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICategorizer_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICategorizer_GetDescription( x, p1, p2 ) \
            (x)->lpVtbl->GetDescription( x, p1, p2 )
        #define ICategorizer_GetCategory( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetCategory( x, p1, p2, p3 )
        #define ICategorizer_GetCategoryInfo( x, p1, p2 ) \
            (x)->lpVtbl->GetCategoryInfo( x, p1, p2 )
        #define ICategorizer_CompareCategory( x, p1, p2, p3 ) \
            (x)->lpVtbl->CompareCategory( x, p1, p2, p3 )
    #endif
    #if (NTDDI_VERSION >= 0x05000000)
        #define IDropTargetHelper_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDropTargetHelper_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDropTargetHelper_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDropTargetHelper_DragEnter( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->DragEnter( x, p1, p2, p3, p4 )
        #define IDropTargetHelper_DragLeave( x ) \
            (x)->lpVtbl->DragLeave( x )
        #define IDropTargetHelper_DragOver( x, p1, p2 ) \
            (x)->lpVtbl->DragOver( x, p1, p2 )
        #define IDropTargetHelper_Drop( x, p1, p2, p3 ) \
            (x)->lpVtbl->Drop( x, p1, p2, p3 )
        #define IDropTargetHelper_Show( x, p ) \
            (x)->lpVtbl->Show( x, p )
        #define IDragSourceHelper_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDragSourceHelper_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDragSourceHelper_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDragSourceHelper_InitializeFromBitmap( x, p1, p2 ) \
            (x)->lpVtbl->InitializeFromBitmap( x, p1, p2 )
        #define IDragSourceHelper_InitializeFromWindow( x, p1, p2, p3 ) \
            (x)->lpVtbl->InitializeFromWindow( x, p1, p2, p3 )
    #endif
    #if (NTDDI_VERSION >= 0x06000000)
        #define IDragSourceHelper2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDragSourceHelper2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDragSourceHelper2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDragSourceHelper2_InitializeFromBitmap( x, p1, p2 ) \
            (x)->lpVtbl->InitializeFromBitmap( x, p1, p2 )
        #define IDragSourceHelper2_InitializeFromWindow( x, p1, p2, p3 ) \
            (x)->lpVtbl->InitializeFromWindow( x, p1, p2, p3 )
        #define IDragSourceHelper2_SetFlags( x, p ) \
            (x)->lpVtbl->SetFlags( x, p )
    #endif
    #define IShellLinkA_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellLinkA_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellLinkA_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellLinkA_GetPath( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetPath( x, p1, p2, p3, p4 )
    #define IShellLinkA_GetIDList( x, p ) \
        (x)->lpVtbl->GetIDList( x, p )
    #define IShellLinkA_SetIDList( x, p ) \
        (x)->lpVtbl->SetIDList( x, p )
    #define IShellLinkA_GetDescription( x, p1, p2 ) \
        (x)->lpVtbl->GetDescription( x, p1, p2 )
    #define IShellLinkA_SetDescription( x, p ) \
        (x)->lpVtbl->SetDescription( x, p )
    #define IShellLinkA_GetWorkingDirectory( x, p1, p2 ) \
        (x)->lpVtbl->GetWorkingDirectory( x, p1, p2 )
    #define IShellLinkA_SetWorkingDirectory( x, p ) \
        (x)->lpVtbl->SetWorkingDirectory( x, p )
    #define IShellLinkA_GetArguments( x, p1, p2 ) \
        (x)->lpVtbl->GetArguments( x, p1, p2 )
    #define IShellLinkA_SetArguments( x, p ) \
        (x)->lpVtbl->SetArguments( x, p )
    #define IShellLinkA_GetHotKey( x, p ) \
        (x)->lpVtbl->GetHotKey( x, p )
    #define IShellLinkA_SetHotKey( x, p ) \
        (x)->lpVtbl->SetHotKey( x, p )
    #define IShellLinkA_GetShowCmd( x, p ) \
        (x)->lpVtbl->GetShowCmd( x, p )
    #define IShellLinkA_SetShowCmd( x, p ) \
        (x)->lpVtbl->SetShowCmd( x, p )
    #define IShellLinkA_GetIconLocation( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetIconLocation( x, p1, p2, p3 )
    #define IShellLinkA_SetIconLocation( x, p1, p2 ) \
        (x)->lpVtbl->SetIconLocation( x, p1, p2 )
    #define IShellLinkA_SetRelativePath( x, p1, p2 ) \
        (x)->lpVtbl->SetRelativePath( x, p1, p2 )
    #define IShellLinkA_Resolve( x, p1, p2 ) \
        (x)->lpVtbl->Resolve( x, p1, p2 )
    #define IShellLinkA_SetPath( x, p ) \
        (x)->lpVtbl->SetPath( x, p )
    #define IShellLinkW_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellLinkW_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellLinkW_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellLinkW_GetPath( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetPath( x, p1, p2, p3, p4 )
    #define IShellLinkW_GetIDList( x, p ) \
        (x)->lpVtbl->GetIDList( x, p )
    #define IShellLinkW_SetIDList( x, p ) \
        (x)->lpVtbl->SetIDList( x, p )
    #define IShellLinkW_GetDescription( x, p1, p2 ) \
        (x)->lpVtbl->GetDescription( x, p1, p2 )
    #define IShellLinkW_SetDescription( x, p ) \
        (x)->lpVtbl->SetDescription( x, p )
    #define IShellLinkW_GetWorkingDirectory( x, p1, p2 ) \
        (x)->lpVtbl->GetWorkingDirectory( x, p1, p2 )
    #define IShellLinkW_SetWorkingDirectory( x, p ) \
        (x)->lpVtbl->SetWorkingDirectory( x, p )
    #define IShellLinkW_GetArguments( x, p1, p2 ) \
        (x)->lpVtbl->GetArguments( x, p1, p2 )
    #define IShellLinkW_SetArguments( x, p ) \
        (x)->lpVtbl->SetArguments( x, p )
    #define IShellLinkW_GetHotKey( x, p ) \
        (x)->lpVtbl->GetHotKey( x, p )
    #define IShellLinkW_SetHotKey( x, p ) \
        (x)->lpVtbl->SetHotKey( x, p )
    #define IShellLinkW_GetShowCmd( x, p ) \
        (x)->lpVtbl->GetShowCmd( x, p )
    #define IShellLinkW_SetShowCmd( x, p ) \
        (x)->lpVtbl->SetShowCmd( x, p )
    #define IShellLinkW_GetIconLocation( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetIconLocation( x, p1, p2, p3 )
    #define IShellLinkW_SetIconLocation( x, p1, p2 ) \
        (x)->lpVtbl->SetIconLocation( x, p1, p2 )
    #define IShellLinkW_SetRelativePath( x, p1, p2 ) \
        (x)->lpVtbl->SetRelativePath( x, p1, p2 )
    #define IShellLinkW_Resolve( x, p1, p2 ) \
        (x)->lpVtbl->Resolve( x, p1, p2 )
    #define IShellLinkW_SetPath( x, p ) \
        (x)->lpVtbl->SetPath( x, p )
    #define IShellLinkDataList_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellLinkDataList_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellLinkDataList_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellLinkDataList_AddDataBlock( x, p ) \
        (x)->lpVtbl->AddDataBlock( x, p )
    #define IShellLinkDataList_CopyDataBlock( x, p1, p2 ) \
        (x)->lpVtbl->CopyDataBlock( x, p1, p2 )
    #define IShellLinkDataList_RemoveDataBlock( x, p ) \
        (x)->lpVtbl->RemoveDataBlock( x, p )
    #define IShellLinkDataList_GetFlags( x, p ) \
        (x)->lpVtbl->GetFlags( x, p )
    #define IShellLinkDataList_SetFlags( x, p ) \
        (x)->lpVtbl->SetFlags( x, p )
    #if (NTDDI_VERSION >= 0x05000000)
        #define IResolveShellLink_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IResolveShellLink_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IResolveShellLink_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IResolveShellLink_ResolveShellLink( x, p1, p2, p3 ) \
            (x)->lpVtbl->ResolveShellLink( x, p1, p2, p3 )
    #endif
    #define IActionProgressDialog_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IActionProgressDialog_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IActionProgressDialog_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IActionProgressDialog_Initialize( x, p1, p2, p3 ) \
        (x)->lpVtbl->Initialize( x, p1, p2, p3 )
    #define IActionProgressDialog_Stop( x ) \
        (x)->lpVtbl->Stop( x )
    #define IHWEventHandler_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IHWEventHandler_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IHWEventHandler_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IHWEventHandler_Initialize( x, p ) \
        (x)->lpVtbl->Initialize( x, p )
    #define IHWEventHandler_HandleEvent( x, p1, p2, p3 ) \
        (x)->lpVtbl->HandleEvent( x, p1, p2, p3 )
    #define IHWEventHandler_HandleEventWithContent( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->HandleEventWithContent( x, p1, p2, p3, p4, p5 )
    #define IHWEventHandler2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IHWEventHandler2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IHWEventHandler2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IHWEventHandler2_Initialize( x, p ) \
        (x)->lpVtbl->Initialize( x, p )
    #define IHWEventHandler2_HandleEvent( x, p1, p2, p3 ) \
        (x)->lpVtbl->HandleEvent( x, p1, p2, p3 )
    #define IHWEventHandler2_HandleEventWithContent( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->HandleEventWithContent( x, p1, p2, p3, p4, p5 )
    #define IHWEventHandler2_HandleEventWithHWND( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->HandleEventWithHWND( x, p1, p2, p3, p4 )
    #define IQueryCancelAutoPlay_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IQueryCancelAutoPlay_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IQueryCancelAutoPlay_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IQueryCancelAutoPlay_AllowAutoPlay( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->AllowAutoPlay( x, p1, p2, p3, p4 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define IDynamicHWHandler_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDynamicHWHandler_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDynamicHWHandler_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDynamicHWHandler_GetDynamicInfo( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetDynamicInfo( x, p1, p2, p3 )
    #endif
    #define IActionProgress_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IActionProgress_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IActionProgress_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IActionProgress_Begin( x, p1, p2 ) \
        (x)->lpVtbl->Begin( x, p1, p2 )
    #define IActionProgress_UpdateProgress( x, p1, p2 ) \
        (x)->lpVtbl->UpdateProgress( x, p1, p2 )
    #define IActionProgress_UpdateText( x, p1, p2, p3 ) \
        (x)->lpVtbl->UpdateText( x, p1, p2, p3 )
    #define IActionProgress_QueryCancel( x, p ) \
        (x)->lpVtbl->QueryCancel( x, p )
    #define IActionProgress_ResetCancel( x ) \
        (x)->lpVtbl->ResetCancel( x )
    #define IActionProgress_End( x ) \
        (x)->lpVtbl->End( x )
    #define IShellExtInit_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellExtInit_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellExtInit_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellExtInit_Initialize( x, p1, p2, p3 ) \
        (x)->lpVtbl->Initialize( x, p1, p2, p3 )
    #define IShellPropSheetExt_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellPropSheetExt_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellPropSheetExt_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellPropSheetExt_AddPages( x, p1, p2 ) \
        (x)->lpVtbl->AddPages( x, p1, p2 )
    #define IShellPropSheetExt_ReplacePage( x, p1, p2, p3 ) \
        (x)->lpVtbl->ReplacePage( x, p1, p2, p3 )
    #define IRemoteComputer_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IRemoteComputer_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IRemoteComputer_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IRemoteComputer_Initialize( x, p1, p2 ) \
        (x)->lpVtbl->Initialize( x, p1, p2 )
    #define IQueryContinue_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IQueryContinue_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IQueryContinue_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IQueryContinue_QueryContinue( x ) \
        (x)->lpVtbl->QueryContinue( x )
    #define IObjectWithCancelEvent_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IObjectWithCancelEvent_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IObjectWithCancelEvent_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IObjectWithCancelEvent_GetCancelEvent( x, p ) \
        (x)->lpVtbl->GetCancelEvent( x, p )
    #define IUserNotification_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IUserNotification_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IUserNotification_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IUserNotification_SetBalloonInfo( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetBalloonInfo( x, p1, p2, p3 )
    #define IUserNotification_SetBalloonRetry( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetBalloonRetry( x, p1, p2, p3 )
    #define IUserNotification_SetIconInfo( x, p1, p2 ) \
        (x)->lpVtbl->SetIconInfo( x, p1, p2 )
    #define IUserNotification_Show( x, p1, p2 ) \
        (x)->lpVtbl->Show( x, p1, p2 )
    #define IUserNotification_PlaySound( x, p ) \
        (x)->lpVtbl->PlaySound( x, p )
    #define IUserNotificationCallback_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IUserNotificationCallback_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IUserNotificationCallback_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IUserNotificationCallback_OnBalloonUserClick( x, p ) \
        (x)->lpVtbl->OnBalloonUserClick( x, p )
    #define IUserNotificationCallback_OnLeftClick( x, p ) \
        (x)->lpVtbl->OnLeftClick( x, p )
    #define IUserNotificationCallback_OnContextMenu( x, p ) \
        (x)->lpVtbl->OnContextMenu( x, p )
    #define IUserNotification2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IUserNotification2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IUserNotification2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IUserNotification2_SetBalloonInfo( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetBalloonInfo( x, p1, p2, p3 )
    #define IUserNotification2_SetBalloonRetry( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetBalloonRetry( x, p1, p2, p3 )
    #define IUserNotification2_SetIconInfo( x, p1, p2 ) \
        (x)->lpVtbl->SetIconInfo( x, p1, p2 )
    #define IUserNotification2_Show( x, p1, p2, p3 ) \
        (x)->lpVtbl->Show( x, p1, p2, p3 )
    #define IUserNotification2_PlaySound( x, p ) \
        (x)->lpVtbl->PlaySound( x, p )
    #define IItemNameLimits_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IItemNameLimits_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IItemNameLimits_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IItemNameLimits_GetValidCharacteristics( x, p1, p2 ) \
        (x)->lpVtbl->GetValidCharacteristics( x, p1, p2 )
    #define IItemNameLimits_GetMaxLength( x, p1, p2 ) \
        (x)->lpVtbl->GetMaxLength( x, p1, p2 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define ISearchFolderItemFactory_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ISearchFolderItemFactory_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ISearchFolderItemFactory_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ISearchFolderItemFactory_SetDisplayName( x, p ) \
            (x)->lpVtbl->SetDisplayName( x, p )
        #define ISearchFolderItemFactory_SetFolderTypeID( x, p ) \
            (x)->lpVtbl->SetFolderTypeID( x, p )
        #define ISearchFolderItemFactory_SetFolderLogicalViewMode( x, p ) \
            (x)->lpVtbl->SetFolderLogicalViewMode( x, p )
        #define ISearchFolderItemFactory_SetIconSize( x, p ) \
            (x)->lpVtbl->SetIconSize( x, p )
        #define ISearchFolderItemFactory_SetVisibleColumns( x, p1, p2 ) \
            (x)->lpVtbl->SetVisibleColumns( x, p1, p2 )
        #define ISearchFolderItemFactory_SetSortColumns( x, p1, p2 ) \
            (x)->lpVtbl->SetSortColumns( x, p1, p2 )
        #define ISearchFolderItemFactory_SetGroupColumn( x, p ) \
            (x)->lpVtbl->SetGroupColumn( x, p )
        #define ISearchFolderItemFactory_SetStacks( x, p1, p2 ) \
            (x)->lpVtbl->SetStacks( x, p1, p2 )
        #define ISearchFolderItemFactory_SetScope( x, p ) \
            (x)->lpVtbl->SetScope( x, p )
        #define ISearchFolderItemFactory_SetCondition( x, p ) \
            (x)->lpVtbl->SetCondition( x, p )
        #define ISearchFolderItemFactory_GetShellItem( x, p1, p2 ) \
            (x)->lpVtbl->GetShellItem( x, p1, p2 )
        #define ISearchFolderItemFactory_GetIDList( x, p ) \
            (x)->lpVtbl->GetIDList( x, p )
    #endif
    #if (_WIN32_IE >= 0x0400)
        #define IExtractImage_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IExtractImage_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IExtractImage_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IExtractImage_GetLocation( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->GetLocation( x, p1, p2, p3, p4, p5, p6 )
        #define IExtractImage_Extract( x, p ) \
            (x)->lpVtbl->Extract( x, p )
    #endif
    #if (_WIN32_IE >= 0x0500)
        #define IExtractImage2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IExtractImage2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IExtractImage2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IExtractImage2_GetLocation( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->GetLocation( x, p1, p2, p3, p4, p5, p6 )
        #define IExtractImage2_Extract( x, p ) \
            (x)->lpVtbl->Extract( x, p )
        #define IExtractImage2_GetDateStamp( x, p ) \
            (x)->lpVtbl->GetDateStamp( x, p )
        #define IThumbnailHandlerFactory_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IThumbnailHandlerFactory_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IThumbnailHandlerFactory_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IThumbnailHandlerFactory_GetThumbnailFactory( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->GetThumbnailFactory( x, p1, p2, p3, p4 )
        #define IParentAndItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IParentAndItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IParentAndItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IParentAndItem_SetParentAndItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetParentAndItem( x, p1, p2, p3 )
        #define IParentAndItem_GetParentAndItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetParentAndItem( x, p1, p2, p3 )
    #endif
    #define IDockingWindow_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IDockingWindow_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IDockingWindow_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IDockingWindow_GetWindow( x, p ) \
        (x)->lpVtbl->GetWindow( x, p )
    #define IDockingWindow_ContextSensitiveHelp( x, p ) \
        (x)->lpVtbl->ContextSensitiveHelp( x, p )
    #define IDockingWindow_ShowDW( x, p ) \
        (x)->lpVtbl->ShowDW( x, p )
    #define IDockingWindow_CloseDW( x, p ) \
        (x)->lpVtbl->CloseDW( x, p )
    #define IDockingWindow_ResizeBorderDW( x, p1, p2, p3 ) \
        (x)->lpVtbl->ResizeBorderDW( x, p1, p2, p3 )
    #define IDeskBand_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IDeskBand_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IDeskBand_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IDeskBand_GetWindow( x, p ) \
        (x)->lpVtbl->GetWindow( x, p )
    #define IDeskBand_ContextSensitiveHelp( x, p ) \
        (x)->lpVtbl->ContextSensitiveHelp( x, p )
    #define IDeskBand_ShowDW( x, p ) \
        (x)->lpVtbl->ShowDW( x, p )
    #define IDeskBand_CloseDW( x, p ) \
        (x)->lpVtbl->CloseDW( x, p )
    #define IDeskBand_ResizeBorderDW( x, p1, p2, p3 ) \
        (x)->lpVtbl->ResizeBorderDW( x, p1, p2, p3 )
    #define IDeskBand_GetBandInfo( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetBandInfo( x, p1, p2, p3 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define IDeskBandInfo_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDeskBandInfo_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDeskBandInfo_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDeskBandInfo_GetDefaultBandWidth( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetDefaultBandWidth( x, p1, p2, p3 )
        #define IDeskBand2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDeskBand2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDeskBand2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDeskBand2_GetWindow( x, p ) \
            (x)->lpVtbl->GetWindow( x, p )
        #define IDeskBand2_ContextSensitiveHelp( x, p ) \
            (x)->lpVtbl->ContextSensitiveHelp( x, p )
        #define IDeskBand2_ShowDW( x, p ) \
            (x)->lpVtbl->ShowDW( x, p )
        #define IDeskBand2_CloseDW( x, p ) \
            (x)->lpVtbl->CloseDW( x, p )
        #define IDeskBand2_ResizeBorderDW( x, p1, p2, p3 ) \
            (x)->lpVtbl->ResizeBorderDW( x, p1, p2, p3 )
        #define IDeskBand2_GetBandInfo( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetBandInfo( x, p1, p2, p3 )
        #define IDeskBand2_CanRenderComposited( x, p ) \
            (x)->lpVtbl->CanRenderComposited( x, p )
        #define IDeskBand2_SetCompositionState( x, p ) \
            (x)->lpVtbl->SetCompositionState( x, p )
        #define IDeskBand2_GetCompositionState( x, p ) \
            (x)->lpVtbl->GetCompositionState( x, p )
    #endif
    #define ITaskbarList_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ITaskbarList_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ITaskbarList_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ITaskbarList_HrInit( x ) \
        (x)->lpVtbl->HrInit( x )
    #define ITaskbarList_AddTab( x, p ) \
        (x)->lpVtbl->AddTab( x, p )
    #define ITaskbarList_DeleteTab( x, p ) \
        (x)->lpVtbl->DeleteTab( x, p )
    #define ITaskbarList_ActivateTab( x, p ) \
        (x)->lpVtbl->ActivateTab( x, p )
    #define ITaskbarList_SetActiveAlt( x, p ) \
        (x)->lpVtbl->SetActiveAlt( x, p )
    #define ITaskbarList2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ITaskbarList2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ITaskbarList2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ITaskbarList2_HrInit( x ) \
        (x)->lpVtbl->HrInit( x )
    #define ITaskbarList2_AddTab( x, p ) \
        (x)->lpVtbl->AddTab( x, p )
    #define ITaskbarList2_DeleteTab( x, p ) \
        (x)->lpVtbl->DeleteTab( x, p )
    #define ITaskbarList2_ActivateTab( x, p ) \
        (x)->lpVtbl->ActivateTab( x, p )
    #define ITaskbarList2_SetActiveAlt( x, p ) \
        (x)->lpVtbl->SetActiveAlt( x, p )
    #define ITaskbarList2_MarkFullscreenWindow( x, p1, p2 ) \
        (x)->lpVtbl->MarkFullscreenWindow( x, p1, p2 )
    #define ITaskbarList3_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ITaskbarList3_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ITaskbarList3_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ITaskbarList3_HrInit( x ) \
        (x)->lpVtbl->HrInit( x )
    #define ITaskbarList3_AddTab( x, p ) \
        (x)->lpVtbl->AddTab( x, p )
    #define ITaskbarList3_DeleteTab( x, p ) \
        (x)->lpVtbl->DeleteTab( x, p )
    #define ITaskbarList3_ActivateTab( x, p ) \
        (x)->lpVtbl->ActivateTab( x, p )
    #define ITaskbarList3_SetActiveAlt( x, p ) \
        (x)->lpVtbl->SetActiveAlt( x, p )
    #define ITaskbarList3_MarkFullscreenWindow( x, p1, p2 ) \
        (x)->lpVtbl->MarkFullscreenWindow( x, p1, p2 )
    #define ITaskbarList3_SetProgressValue( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetProgressValue( x, p1, p2, p3 )
    #define ITaskbarList3_SetProgressState( x, p1, p2 ) \
        (x)->lpVtbl->SetProgressState( x, p1, p2 )
    #define ITaskbarList3_RegisterTab( x, p1, p2 ) \
        (x)->lpVtbl->RegisterTab( x, p1, p2 )
    #define ITaskbarList3_UnregisterTab( x, p ) \
        (x)->lpVtbl->UnregisterTab( x, p )
    #define ITaskbarList3_SetTabOrder( x, p1, p2 ) \
        (x)->lpVtbl->SetTabOrder( x, p1, p2 )
    #define ITaskbarList3_SetTabActive( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetTabActive( x, p1, p2, p3 )
    #define ITaskbarList3_ThumbBarAddButtons( x, p1, p2, p3 ) \
        (x)->lpVtbl->ThumbBarAddButtons( x, p1, p2, p3 )
    #define ITaskbarList3_ThumbBarUpdateButtons( x, p1, p2, p3 ) \
        (x)->lpVtbl->ThumbBarUpdateButtons( x, p1, p2, p3 )
    #define ITaskbarList3_ThumbBarSetImageList( x, p1, p2 ) \
        (x)->lpVtbl->ThumbBarSetImageList( x, p1, p2 )
    #define ITaskbarList3_SetOverlayIcon( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetOverlayIcon( x, p1, p2, p3 )
    #define ITaskbarList3_SetThumbnailTooltip( x, p1, p2 ) \
        (x)->lpVtbl->SetThumbnailTooltip( x, p1, p2 )
    #define ITaskbarList3_SetThumbnailClip( x, p1, p2 ) \
        (x)->lpVtbl->SetThumbnailClip( x, p1, p2 )
    #define ITaskbarList4_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ITaskbarList4_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ITaskbarList4_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ITaskbarList4_HrInit( x ) \
        (x)->lpVtbl->HrInit( x )
    #define ITaskbarList4_AddTab( x, p ) \
        (x)->lpVtbl->AddTab( x, p )
    #define ITaskbarList4_DeleteTab( x, p ) \
        (x)->lpVtbl->DeleteTab( x, p )
    #define ITaskbarList4_ActivateTab( x, p ) \
        (x)->lpVtbl->ActivateTab( x, p )
    #define ITaskbarList4_SetActiveAlt( x, p ) \
        (x)->lpVtbl->SetActiveAlt( x, p )
    #define ITaskbarList4_MarkFullscreenWindow( x, p1, p2 ) \
        (x)->lpVtbl->MarkFullscreenWindow( x, p1, p2 )
    #define ITaskbarList4_SetProgressValue( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetProgressValue( x, p1, p2, p3 )
    #define ITaskbarList4_SetProgressState( x, p1, p2 ) \
        (x)->lpVtbl->SetProgressState( x, p1, p2 )
    #define ITaskbarList4_RegisterTab( x, p1, p2 ) \
        (x)->lpVtbl->RegisterTab( x, p1, p2 )
    #define ITaskbarList4_UnregisterTab( x, p ) \
        (x)->lpVtbl->UnregisterTab( x, p )
    #define ITaskbarList4_SetTabOrder( x, p1, p2 ) \
        (x)->lpVtbl->SetTabOrder( x, p1, p2 )
    #define ITaskbarList4_SetTabActive( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetTabActive( x, p1, p2, p3 )
    #define ITaskbarList4_ThumbBarAddButtons( x, p1, p2, p3 ) \
        (x)->lpVtbl->ThumbBarAddButtons( x, p1, p2, p3 )
    #define ITaskbarList4_ThumbBarUpdateButtons( x, p1, p2, p3 ) \
        (x)->lpVtbl->ThumbBarUpdateButtons( x, p1, p2, p3 )
    #define ITaskbarList4_ThumbBarSetImageList( x, p1, p2 ) \
        (x)->lpVtbl->ThumbBarSetImageList( x, p1, p2 )
    #define ITaskbarList4_SetOverlayIcon( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetOverlayIcon( x, p1, p2, p3 )
    #define ITaskbarList4_SetThumbnailTooltip( x, p1, p2 ) \
        (x)->lpVtbl->SetThumbnailTooltip( x, p1, p2 )
    #define ITaskbarList4_SetThumbnailClip( x, p1, p2 ) \
        (x)->lpVtbl->SetThumbnailClip( x, p1, p2 )
    #define ITaskbarList4_SetTabProperties( x, p1, p2 ) \
        (x)->lpVtbl->SetTabProperties( x, p1, p2 )
    #define IStartMenuPinnedList_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IStartMenuPinnedList_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IStartMenuPinnedList_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IStartMenuPinnedList_RemoveFromList( x, p ) \
        (x)->lpVtbl->RemoveFromList( x, p )
    #define ICDBurn_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ICDBurn_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ICDBurn_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ICDBurn_GetRecorderDriveLetter( x, p1, p2 ) \
        (x)->lpVtbl->GetRecorderDriveLetter( x, p1, p2 )
    #define ICDBurn_Burn( x, p ) \
        (x)->lpVtbl->Burn( x, p )
    #define ICDBurn_HasRecordableDrive( x, p ) \
        (x)->lpVtbl->HasRecordableDrive( x, p )
    #define IWizardSite_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IWizardSite_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IWizardSite_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IWizardSite_GetPreviousPage( x, p ) \
        (x)->lpVtbl->GetPreviousPage( x, p )
    #define IWizardSite_GetNextPage( x, p ) \
        (x)->lpVtbl->GetNextPage( x, p )
    #define IWizardSite_GetCancelledPage( x, p ) \
        (x)->lpVtbl->GetCancelledPage( x, p )
    #define IWizardExtension_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IWizardExtension_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IWizardExtension_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IWizardExtension_AddPages( x, p1, p2, p3 ) \
        (x)->lpVtbl->AddPages( x, p1, p2, p3 )
    #define IWizardExtension_GetFirstPage( x, p ) \
        (x)->lpVtbl->GetFirstPage( x, p )
    #define IWizardExtension_GetLastPage( x, p ) \
        (x)->lpVtbl->GetLastPage( x, p )
    #define IWebWizardExtension_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IWebWizardExtension_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IWebWizardExtension_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IWebWizardExtension_AddPages( x, p1, p2, p3 ) \
        (x)->lpVtbl->AddPages( x, p1, p2, p3 )
    #define IWebWizardExtension_GetFirstPage( x, p ) \
        (x)->lpVtbl->GetFirstPage( x, p )
    #define IWebWizardExtension_GetLastPage( x, p ) \
        (x)->lpVtbl->GetLastPage( x, p )
    #define IWebWizardExtension_SetInitialURL( x, p ) \
        (x)->lpVtbl->SetInitialURL( x, p )
    #define IWebWizardExtension_SetErrorURL( x, p ) \
        (x)->lpVtbl->SetErrorURL( x, p )
    #define IPublishingWizard_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IPublishingWizard_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IPublishingWizard_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IPublishingWizard_AddPages( x, p1, p2, p3 ) \
        (x)->lpVtbl->AddPages( x, p1, p2, p3 )
    #define IPublishingWizard_GetFirstPage( x, p ) \
        (x)->lpVtbl->GetFirstPage( x, p )
    #define IPublishingWizard_GetLastPage( x, p ) \
        (x)->lpVtbl->GetLastPage( x, p )
    #define IPublishingWizard_Initialize( x, p1, p2, p3 ) \
        (x)->lpVtbl->Initialize( x, p1, p2, p3 )
    #define IPublishingWizard_GetTransferManifest( x, p1, p2 ) \
        (x)->lpVtbl->GetTransferManifest( x, p1, p2 )
    #if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
        #define IFolderViewHost_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFolderViewHost_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFolderViewHost_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFolderViewHost_Initialize( x, p1, p2, p3 ) \
            (x)->lpVtbl->Initialize( x, p1, p2, p3 )
    #endif
    #if (_WIN32_IE >= 0x0700)
        #define IExplorerBrowserEvents_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IExplorerBrowserEvents_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IExplorerBrowserEvents_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IExplorerBrowserEvents_OnNavigationPending( x, p ) \
            (x)->lpVtbl->OnNavigationPending( x, p )
        #define IExplorerBrowserEvents_OnViewCreated( x, p ) \
            (x)->lpVtbl->OnViewCreated( x, p )
        #define IExplorerBrowserEvents_OnNavigationCompleted( x, p ) \
            (x)->lpVtbl->OnNavigationCompleted( x, p )
        #define IExplorerBrowserEvents_OnNavigationFailed( x, p ) \
            (x)->lpVtbl->OnNavigationFailed( x, p )
        #define IExplorerBrowser_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IExplorerBrowser_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IExplorerBrowser_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IExplorerBrowser_Initialize( x, p1, p2, p3 ) \
            (x)->lpVtbl->Initialize( x, p1, p2, p3 )
        #define IExplorerBrowser_Destroy( x ) \
            (x)->lpVtbl->Destroy( x )
        #define IExplorerBrowser_SetRect( x, p1, p2 ) \
            (x)->lpVtbl->SetRect( x, p1, p2 )
        #define IExplorerBrowser_SetPropertyBag( x, p ) \
            (x)->lpVtbl->SetPropertyBag( x, p )
        #define IExplorerBrowser_SetEmptyText( x, p ) \
            (x)->lpVtbl->SetEmptyText( x, p )
        #define IExplorerBrowser_SetFolderSettings( x, p ) \
            (x)->lpVtbl->SetFolderSettings( x, p )
        #define IExplorerBrowser_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define IExplorerBrowser_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define IExplorerBrowser_SetOptions( x, p ) \
            (x)->lpVtbl->SetOptions( x, p )
        #define IExplorerBrowser_GetOptions( x, p ) \
            (x)->lpVtbl->GetOptions( x, p )
        #define IExplorerBrowser_BrowseToIDList( x, p1, p2 ) \
            (x)->lpVtbl->BrowseToIDList( x, p1, p2 )
        #define IExplorerBrowser_BrowseToObject( x, p1, p2 ) \
            (x)->lpVtbl->BrowseToObject( x, p1, p2 )
        #define IExplorerBrowser_FillFromObject( x, p1, p2 ) \
            (x)->lpVtbl->FillFromObject( x, p1, p2 )
        #define IExplorerBrowser_RemoveAll( x ) \
            (x)->lpVtbl->RemoveAll( x )
        #define IExplorerBrowser_GetCurrentView( x, p1, p2 ) \
            (x)->lpVtbl->GetCurrentView( x, p1, p2 )
        #define IAccessibleObject_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IAccessibleObject_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IAccessibleObject_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IAccessibleObject_SetAccessibleName( x, p ) \
            (x)->lpVtbl->SetAccessibleName( x, p )
    #endif
    #if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
        #define IResultsFolder_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IResultsFolder_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IResultsFolder_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IResultsFolder_AddItem( x, p ) \
            (x)->lpVtbl->AddItem( x, p )
        #define IResultsFolder_AddIDList( x, p1, p2 ) \
            (x)->lpVtbl->AddIDList( x, p1, p2 )
        #define IResultsFolder_RemoveItem( x, p ) \
            (x)->lpVtbl->RemoveItem( x, p )
        #define IResultsFolder_RemoveIDList( x, p ) \
            (x)->lpVtbl->RemoveIDList( x, p )
        #define IResultsFolder_RemoveAll( x ) \
            (x)->lpVtbl->RemoveAll( x )
    #endif
    #if (_WIN32_IE >= 0x0700)
        #define IEnumObjects_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IEnumObjects_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IEnumObjects_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IEnumObjects_Next( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->Next( x, p1, p2, p3, p4 )
        #define IEnumObjects_Skip( x, p ) \
            (x)->lpVtbl->Skip( x, p )
        #define IEnumObjects_Reset( x ) \
            (x)->lpVtbl->Reset( x )
        #define IEnumObjects_Clone( x, p ) \
            (x)->lpVtbl->Clone( x, p )
        #define IOperationsProgressDialog_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IOperationsProgressDialog_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IOperationsProgressDialog_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IOperationsProgressDialog_StartProgressDialog( x, p1, p2 ) \
            (x)->lpVtbl->StartProgressDialog( x, p1, p2 )
        #define IOperationsProgressDialog_StopProgressDialog( x ) \
            (x)->lpVtbl->StopProgressDialog( x )
        #define IOperationsProgressDialog_SetOperation( x, p ) \
            (x)->lpVtbl->SetOperation( x, p )
        #define IOperationsProgressDialog_SetMode( x, p ) \
            (x)->lpVtbl->SetMode( x, p )
        #define IOperationsProgressDialog_UpdateProgress( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->UpdateProgress( x, p1, p2, p3, p4, p5, p6 )
        #define IOperationsProgressDialog_UpdateLocations( x, p1, p2, p3 ) \
            (x)->lpVtbl->UpdateLocations( x, p1, p2, p3 )
        #define IOperationsProgressDialog_ResetTimer( x ) \
            (x)->lpVtbl->ResetTimer( x )
        #define IOperationsProgressDialog_PauseTimer( x ) \
            (x)->lpVtbl->PauseTimer( x )
        #define IOperationsProgressDialog_ResumeTimer( x ) \
            (x)->lpVtbl->ResumeTimer( x )
        #define IOperationsProgressDialog_GetMilliseconds( x, p1, p2 ) \
            (x)->lpVtbl->GetMilliseconds( x, p1, p2 )
        #define IOperationsProgressDialog_GetOperationStatus( x, p ) \
            (x)->lpVtbl->GetOperationStatus( x, p )
        #define IIOCancelInformation_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IIOCancelInformation_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IIOCancelInformation_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IIOCancelInformation_SetCancelInformation( x, p1, p2 ) \
            (x)->lpVtbl->SetCancelInformation( x, p1, p2 )
        #define IIOCancelInformation_GetCancelInformation( x, p1, p2 ) \
            (x)->lpVtbl->GetCancelInformation( x, p1, p2 )
        #define IFileOperation_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileOperation_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileOperation_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileOperation_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define IFileOperation_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define IFileOperation_SetOperationFlags( x, p ) \
            (x)->lpVtbl->SetOperationFlags( x, p )
        #define IFileOperation_SetProgressMessage( x, p ) \
            (x)->lpVtbl->SetProgressMessage( x, p )
        #define IFileOperation_SetProgressDialog( x, p ) \
            (x)->lpVtbl->SetProgressDialog( x, p )
        #define IFileOperation_SetProperties( x, p ) \
            (x)->lpVtbl->SetProperties( x, p )
        #define IFileOperation_SetOwnerWindow( x, p ) \
            (x)->lpVtbl->SetOwnerWindow( x, p )
        #define IFileOperation_ApplyPropertiesToItem( x, p ) \
            (x)->lpVtbl->ApplyPropertiesToItem( x, p )
        #define IFileOperation_ApplyPropertiesToItems( x, p ) \
            (x)->lpVtbl->ApplyPropertiesToItems( x, p )
        #define IFileOperation_RenameItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->RenameItem( x, p1, p2, p3 )
        #define IFileOperation_RenameItems( x, p1, p2 ) \
            (x)->lpVtbl->RenameItems( x, p1, p2 )
        #define IFileOperation_MoveItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->MoveItem( x, p1, p2, p3, p4 )
        #define IFileOperation_MoveItems( x, p1, p2 ) \
            (x)->lpVtbl->MoveItems( x, p1, p2 )
        #define IFileOperation_CopyItem( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->CopyItem( x, p1, p2, p3, p4 )
        #define IFileOperation_CopyItems( x, p1, p2 ) \
            (x)->lpVtbl->CopyItems( x, p1, p2 )
        #define IFileOperation_DeleteItem( x, p1, p2 ) \
            (x)->lpVtbl->DeleteItem( x, p1, p2 )
        #define IFileOperation_DeleteItems( x, p ) \
            (x)->lpVtbl->DeleteItems( x, p )
        #define IFileOperation_NewItem( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->NewItem( x, p1, p2, p3, p4, p5 )
        #define IFileOperation_PerformOperations( x ) \
            (x)->lpVtbl->PerformOperations( x )
        #define IFileOperation_GetAnyOperationsAborted( x, p ) \
            (x)->lpVtbl->GetAnyOperationsAborted( x, p )
        #define IObjectProvider_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IObjectProvider_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IObjectProvider_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IObjectProvider_QueryObject( x, p1, p2, p3 ) \
            (x)->lpVtbl->QueryObject( x, p1, p2, p3 )
   #endif
   #if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
        #define INamespaceWalkCB_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define INamespaceWalkCB_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define INamespaceWalkCB_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define INamespaceWalkCB_FoundItem( x, p1, p2 ) \
            (x)->lpVtbl->FoundItem( x, p1, p2 )
        #define INamespaceWalkCB_EnterFolder( x, p1, p2 ) \
            (x)->lpVtbl->EnterFolder( x, p1, p2 )
        #define INamespaceWalkCB_LeaveFolder( x, p1, p2 ) \
            (x)->lpVtbl->LeaveFolder( x, p1, p2 )
        #define INamespaceWalkCB_InitializeProgressDialog( x, p1, p2 ) \
            (x)->lpVtbl->InitializeProgressDialog( x, p1, p2 )
    #endif
    #if (_WIN32_IE >= 0x0700)
        #define INamespaceWalkCB2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define INamespaceWalkCB2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define INamespaceWalkCB2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define INamespaceWalkCB2_FoundItem( x, p1, p2 ) \
            (x)->lpVtbl->FoundItem( x, p1, p2 )
        #define INamespaceWalkCB2_EnterFolder( x, p1, p2 ) \
            (x)->lpVtbl->EnterFolder( x, p1, p2 )
        #define INamespaceWalkCB2_LeaveFolder( x, p1, p2 ) \
            (x)->lpVtbl->LeaveFolder( x, p1, p2 )
        #define INamespaceWalkCB2_InitializeProgressDialog( x, p1, p2 ) \
            (x)->lpVtbl->InitializeProgressDialog( x, p1, p2 )
        #define INamespaceWalkCB2_WalkComplete( x, p ) \
            (x)->lpVtbl->WalkComplete( x, p )
    #endif
    #if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
        #define INamespaceWalk_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define INamespaceWalk_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define INamespaceWalk_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define INamespaceWalk_Walk( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->Walk( x, p1, p2, p3, p4 )
        #define INamespaceWalk_GetIDArrayResult( x, p1, p2 ) \
            (x)->lpVtbl->GetIDArrayResult( x, p1, p2 )
    #endif
    #define IAutoCompleteDropDown_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IAutoCompleteDropDown_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IAutoCompleteDropDown_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IAutoCompleteDropDown_GetDropDownStatus( x, p1, p2 ) \
        (x)->lpVtbl->GetDropDownStatus( x, p1, p2 )
    #define IAutoCompleteDropDown_ResetEnumerator( x ) \
        (x)->lpVtbl->ResetEnumerator( x )
    #if (_WIN32_IE >= 0x0400)
        #define IBandSite_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IBandSite_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IBandSite_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IBandSite_AddBand( x, p ) \
            (x)->lpVtbl->AddBand( x, p )
        #define IBandSite_EnumBands( x, p1, p2 ) \
            (x)->lpVtbl->EnumBands( x, p1, p2 )
        #define IBandSite_QueryBand( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->QueryBand( x, p1, p2, p3, p4, p5 )
        #define IBandSite_SetBandState( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetBandState( x, p1, p2, p3 )
        #define IBandSite_RemoveBand( x, p ) \
            (x)->lpVtbl->RemoveBand( x, p )
        #define IBandSite_GetBandObject( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetBandObject( x, p1, p2, p3 )
        #define IBandSite_SetBandSiteInfo( x, p ) \
            (x)->lpVtbl->SetBandSiteInfo( x, p )
        #define IBandSite_GetBandSiteInfo( x, p ) \
            (x)->lpVtbl->GetBandSiteInfo( x, p )
    #endif
    #if (NTDDI_VERSION >= 0x05010000)
        #define IModalWindow_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IModalWindow_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IModalWindow_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IModalWindow_Show( x, p ) \
            (x)->lpVtbl->Show( x, p )
        #define ICDBurnExt_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICDBurnExt_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICDBurnExt_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICDBurnExt_GetSupportedActionTypes( x, p ) \
            (x)->lpVtbl->GetSupportedActionTypes( x, p )
    #endif
    #define IContextMenuSite_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IContextMenuSite_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IContextMenuSite_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IContextMenuSite_DoContextMenuPopup( x, p1, p2, p3 ) \
        (x)->lpVtbl->DoContextMenuPopup( x, p1, p2, p3 )
    #define IEnumReadyCallback_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumReadyCallback_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumReadyCallback_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumReadyCallback_EnumReady( x ) \
        (x)->lpVtbl->EnumReady( x )
    #define IEnumerableView_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumerableView_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumerableView_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumerableView_SetEnumReadyCallback( x, p ) \
        (x)->lpVtbl->SetEnumReadyCallback( x, p )
    #define IEnumerableView_CreateEnumIDListFromContents( x, p1, p2, p3 ) \
        (x)->lpVtbl->CreateEnumIDListFromContents( x, p1, p2, p3 )
    #if (NTDDI_VERSION >= 0x05010000) || (_WIN32_IE >= 0x0700)
        #define IInsertItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IInsertItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IInsertItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IInsertItem_InsertItem( x, p ) \
            (x)->lpVtbl->InsertItem( x, p )
    #endif
    #if (NTDDI_VERSION >= 0x05010000)
        #define IMenuBand_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IMenuBand_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IMenuBand_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IMenuBand_IsMenuMessage( x, p ) \
            (x)->lpVtbl->IsMenuMessage( x, p )
        #define IMenuBand_TranslateMenuMessage( x, p1, p2 ) \
            (x)->lpVtbl->TranslateMenuMessage( x, p1, p2 )
        #define IFolderBandPriv_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFolderBandPriv_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFolderBandPriv_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFolderBandPriv_SetCascade( x, p ) \
            (x)->lpVtbl->SetCascade( x, p )
        #define IFolderBandPriv_SetAccelerators( x, p ) \
            (x)->lpVtbl->SetAccelerators( x, p )
        #define IFolderBandPriv_SetNoIcons( x, p ) \
            (x)->lpVtbl->SetNoIcons( x, p )
        #define IFolderBandPriv_SetNoText( x, p ) \
            (x)->lpVtbl->SetNoText( x, p )
        #define IRegTreeItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IRegTreeItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IRegTreeItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IRegTreeItem_GetCheckState( x, p ) \
            (x)->lpVtbl->GetCheckState( x, p )
        #define IRegTreeItem_SetCheckState( x, p ) \
            (x)->lpVtbl->SetCheckState( x, p )
        #define IImageRecompress_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IImageRecompress_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IImageRecompress_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IImageRecompress_RecompressImage( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->RecompressImage( x, p1, p2, p3, p4, p5, p6 )
    #endif
    #if (_WIN32_IE >= 0x0600)
        #define IDeskBar_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDeskBar_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDeskBar_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDeskBar_GetWindow( x, p ) \
            (x)->lpVtbl->GetWindow( x, p )
        #define IDeskBar_ContextSensitiveHelp( x, p ) \
            (x)->lpVtbl->ContextSensitiveHelp( x, p )
        #define IDeskBar_SetClient( x, p ) \
            (x)->lpVtbl->SetClient( x, p )
        #define IDeskBar_GetClient( x, p ) \
            (x)->lpVtbl->GetClient( x, p )
        #define IDeskBar_OnPosRectChangeDB( x, p ) \
            (x)->lpVtbl->OnPosRectChangeDB( x, p )
        #define IMenuPopup_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IMenuPopup_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IMenuPopup_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IMenuPopup_GetWindow( x, p ) \
            (x)->lpVtbl->GetWindow( x, p )
        #define IMenuPopup_ContextSensitiveHelp( x, p ) \
            (x)->lpVtbl->ContextSensitiveHelp( x, p )
        #define IMenuPopup_SetClient( x, p ) \
            (x)->lpVtbl->SetClient( x, p )
        #define IMenuPopup_GetClient( x, p ) \
            (x)->lpVtbl->GetClient( x, p )
        #define IMenuPopup_OnPosRectChangeDB( x, p ) \
            (x)->lpVtbl->OnPosRectChangeDB( x, p )
        #define IMenuPopup_Popup( x, p1, p2, p3 ) \
            (x)->lpVtbl->Popup( x, p1, p2, p3 )
        #define IMenuPopup_OnSelect( x, p ) \
            (x)->lpVtbl->OnSelect( x, p )
        #define IMenuPopup_SetSubMenu( x, p1, p2 ) \
            (x)->lpVtbl->SetSubMenu( x, p1, p2 )
    #endif
    #if (NTDDI_VERSION >= 0x06000000)
        #define IFileIsInUse_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileIsInUse_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileIsInUse_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileIsInUse_GetAppName( x, p ) \
            (x)->lpVtbl->GetAppName( x, p )
        #define IFileIsInUse_GetUsage( x, p ) \
            (x)->lpVtbl->GetUsage( x, p )
        #define IFileIsInUse_GetCapabilities( x, p ) \
            (x)->lpVtbl->GetCapabilities( x, p )
        #define IFileIsInUse_GetSwitchToHWND( x, p ) \
            (x)->lpVtbl->GetSwitchToHWND( x, p )
        #define IFileIsInUse_CloseFile( x ) \
            (x)->lpVtbl->CloseFile( x )
        #define IFileDialogEvents_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileDialogEvents_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileDialogEvents_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileDialogEvents_OnFileOk( x, p ) \
            (x)->lpVtbl->OnFileOk( x, p )
        #define IFileDialogEvents_OnFolderChanging( x, p1, p2 ) \
            (x)->lpVtbl->OnFolderChanging( x, p1, p2 )
        #define IFileDialogEvents_OnFolderChange( x, p ) \
            (x)->lpVtbl->OnFolderChange( x, p )
        #define IFileDialogEvents_OnSelectionChange( x, p ) \
            (x)->lpVtbl->OnSelectionChange( x, p )
        #define IFileDialogEvents_OnShareViolation( x, p1, p2, p3 ) \
            (x)->lpVtbl->OnShareViolation( x, p1, p2, p3 )
        #define IFileDialogEvents_OnTypeChange( x, p ) \
            (x)->lpVtbl->OnTypeChange( x, p )
        #define IFileDialogEvents_OnOverwrite( x, p1, p2, p3 ) \
            (x)->lpVtbl->OnOverwrite( x, p1, p2, p3 )
        #define IFileDialog_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileDialog_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileDialog_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileDialog_Show( x, p ) \
            (x)->lpVtbl->Show( x, p )
        #define IFileDialog_SetFileTypes( x, p1, p2 ) \
            (x)->lpVtbl->SetFileTypes( x, p1, p2 )
        #define IFileDialog_SetFileTypeIndex( x, p ) \
            (x)->lpVtbl->SetFileTypeIndex( x, p )
        #define IFileDialog_GetFileTypeIndex( x, p ) \
            (x)->lpVtbl->GetFileTypeIndex( x, p )
        #define IFileDialog_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define IFileDialog_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define IFileDialog_SetOptions( x, p ) \
            (x)->lpVtbl->SetOptions( x, p )
        #define IFileDialog_GetOptions( x, p ) \
            (x)->lpVtbl->GetOptions( x, p )
        #define IFileDialog_SetDefaultFolder( x, p ) \
            (x)->lpVtbl->SetDefaultFolder( x, p )
        #define IFileDialog_SetFolder( x, p ) \
            (x)->lpVtbl->SetFolder( x, p )
        #define IFileDialog_GetFolder( x, p ) \
            (x)->lpVtbl->GetFolder( x, p )
        #define IFileDialog_GetCurrentSelection( x, p ) \
            (x)->lpVtbl->GetCurrentSelection( x, p )
        #define IFileDialog_SetFileName( x, p ) \
            (x)->lpVtbl->SetFileName( x, p )
        #define IFileDialog_GetFileName( x, p ) \
            (x)->lpVtbl->GetFileName( x, p )
        #define IFileDialog_SetTitle( x, p ) \
            (x)->lpVtbl->SetTitle( x, p )
        #define IFileDialog_SetOkButtonLabel( x, p ) \
            (x)->lpVtbl->SetOkButtonLabel( x, p )
        #define IFileDialog_SetFileNameLabel( x, p ) \
            (x)->lpVtbl->SetFileNameLabel( x, p )
        #define IFileDialog_GetResult( x, p ) \
            (x)->lpVtbl->GetResult( x, p )
        #define IFileDialog_AddPlace( x, p1, p2 ) \
            (x)->lpVtbl->AddPlace( x, p1, p2 )
        #define IFileDialog_SetDefaultExtension( x, p ) \
            (x)->lpVtbl->SetDefaultExtension( x, p )
        #define IFileDialog_Close( x, p ) \
            (x)->lpVtbl->Close( x, p )
        #define IFileDialog_SetClientGuid( x, p ) \
            (x)->lpVtbl->SetClientGuid( x, p )
        #define IFileDialog_ClearClientData( x ) \
            (x)->lpVtbl->ClearClientData( x )
        #define IFileDialog_SetFilter( x, p ) \
            (x)->lpVtbl->SetFilter( x, p )
        #define IFileSaveDialog_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileSaveDialog_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileSaveDialog_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileSaveDialog_Show( x, p ) \
            (x)->lpVtbl->Show( x, p )
        #define IFileSaveDialog_SetFileTypes( x, p1, p2 ) \
            (x)->lpVtbl->SetFileTypes( x, p1, p2 )
        #define IFileSaveDialog_SetFileTypeIndex( x, p ) \
            (x)->lpVtbl->SetFileTypeIndex( x, p )
        #define IFileSaveDialog_GetFileTypeIndex( x, p ) \
            (x)->lpVtbl->GetFileTypeIndex( x, p )
        #define IFileSaveDialog_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define IFileSaveDialog_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define IFileSaveDialog_SetOptions( x, p ) \
            (x)->lpVtbl->SetOptions( x, p )
        #define IFileSaveDialog_GetOptions( x, p ) \
            (x)->lpVtbl->GetOptions( x, p )
        #define IFileSaveDialog_SetDefaultFolder( x, p ) \
            (x)->lpVtbl->SetDefaultFolder( x, p )
        #define IFileSaveDialog_SetFolder( x, p ) \
            (x)->lpVtbl->SetFolder( x, p )
        #define IFileSaveDialog_GetFolder( x, p ) \
            (x)->lpVtbl->GetFolder( x, p )
        #define IFileSaveDialog_GetCurrentSelection( x, p ) \
            (x)->lpVtbl->GetCurrentSelection( x, p )
        #define IFileSaveDialog_SetFileName( x, p ) \
            (x)->lpVtbl->SetFileName( x, p )
        #define IFileSaveDialog_GetFileName( x, p ) \
            (x)->lpVtbl->GetFileName( x, p )
        #define IFileSaveDialog_SetTitle( x, p ) \
            (x)->lpVtbl->SetTitle( x, p )
        #define IFileSaveDialog_SetOkButtonLabel( x, p ) \
            (x)->lpVtbl->SetOkButtonLabel( x, p )
        #define IFileSaveDialog_SetFileNameLabel( x, p ) \
            (x)->lpVtbl->SetFileNameLabel( x, p )
        #define IFileSaveDialog_GetResult( x, p ) \
            (x)->lpVtbl->GetResult( x, p )
        #define IFileSaveDialog_AddPlace( x, p1, p2 ) \
            (x)->lpVtbl->AddPlace( x, p1, p2 )
        #define IFileSaveDialog_SetDefaultExtension( x, p ) \
            (x)->lpVtbl->SetDefaultExtension( x, p )
        #define IFileSaveDialog_Close( x, p ) \
            (x)->lpVtbl->Close( x, p )
        #define IFileSaveDialog_SetClientGuid( x, p ) \
            (x)->lpVtbl->SetClientGuid( x, p )
        #define IFileSaveDialog_ClearClientData( x ) \
            (x)->lpVtbl->ClearClientData( x )
        #define IFileSaveDialog_SetFilter( x, p ) \
            (x)->lpVtbl->SetFilter( x, p )
        #define IFileSaveDialog_SetSaveAsItem( x, p ) \
            (x)->lpVtbl->SetSaveAsItem( x, p )
        #define IFileSaveDialog_SetProperties( x, p ) \
            (x)->lpVtbl->SetProperties( x, p )
        #define IFileSaveDialog_SetCollectedProperties( x, p1, p2 ) \
            (x)->lpVtbl->SetCollectedProperties( x, p1, p2 )
        #define IFileSaveDialog_GetProperties( x, p ) \
            (x)->lpVtbl->GetProperties( x, p )
        #define IFileSaveDialog_ApplyProperties( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->ApplyProperties( x, p1, p2, p3, p4 )
        #define IFileOpenDialog_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileOpenDialog_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileOpenDialog_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileOpenDialog_Show( x, p ) \
            (x)->lpVtbl->Show( x, p )
        #define IFileOpenDialog_SetFileTypes( x, p1, p2 ) \
            (x)->lpVtbl->SetFileTypes( x, p1, p2 )
        #define IFileOpenDialog_SetFileTypeIndex( x, p ) \
            (x)->lpVtbl->SetFileTypeIndex( x, p )
        #define IFileOpenDialog_GetFileTypeIndex( x, p ) \
            (x)->lpVtbl->GetFileTypeIndex( x, p )
        #define IFileOpenDialog_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define IFileOpenDialog_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define IFileOpenDialog_SetOptions( x, p ) \
            (x)->lpVtbl->SetOptions( x, p )
        #define IFileOpenDialog_GetOptions( x, p ) \
            (x)->lpVtbl->GetOptions( x, p )
        #define IFileOpenDialog_SetDefaultFolder( x, p ) \
            (x)->lpVtbl->SetDefaultFolder( x, p )
        #define IFileOpenDialog_SetFolder( x, p ) \
            (x)->lpVtbl->SetFolder( x, p )
        #define IFileOpenDialog_GetFolder( x, p ) \
            (x)->lpVtbl->GetFolder( x, p )
        #define IFileOpenDialog_GetCurrentSelection( x, p ) \
            (x)->lpVtbl->GetCurrentSelection( x, p )
        #define IFileOpenDialog_SetFileName( x, p ) \
            (x)->lpVtbl->SetFileName( x, p )
        #define IFileOpenDialog_GetFileName( x, p ) \
            (x)->lpVtbl->GetFileName( x, p )
        #define IFileOpenDialog_SetTitle( x, p ) \
            (x)->lpVtbl->SetTitle( x, p )
        #define IFileOpenDialog_SetOkButtonLabel( x, p ) \
            (x)->lpVtbl->SetOkButtonLabel( x, p )
        #define IFileOpenDialog_SetFileNameLabel( x, p ) \
            (x)->lpVtbl->SetFileNameLabel( x, p )
        #define IFileOpenDialog_GetResult( x, p ) \
            (x)->lpVtbl->GetResult( x, p )
        #define IFileOpenDialog_AddPlace( x, p1, p2 ) \
            (x)->lpVtbl->AddPlace( x, p1, p2 )
        #define IFileOpenDialog_SetDefaultExtension( x, p ) \
            (x)->lpVtbl->SetDefaultExtension( x, p )
        #define IFileOpenDialog_Close( x, p ) \
            (x)->lpVtbl->Close( x, p )
        #define IFileOpenDialog_SetClientGuid( x, p ) \
            (x)->lpVtbl->SetClientGuid( x, p )
        #define IFileOpenDialog_ClearClientData( x ) \
            (x)->lpVtbl->ClearClientData( x )
        #define IFileOpenDialog_SetFilter( x, p ) \
            (x)->lpVtbl->SetFilter( x, p )
        #define IFileOpenDialog_GetResults( x, p ) \
            (x)->lpVtbl->GetResults( x, p )
        #define IFileOpenDialog_GetSelectedItems( x, p ) \
            (x)->lpVtbl->GetSelectedItems( x, p )
        #define IFileDialogCustomize_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileDialogCustomize_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileDialogCustomize_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileDialogCustomize_EnableOpenDropDown( x, p ) \
            (x)->lpVtbl->EnableOpenDropDown( x, p )
        #define IFileDialogCustomize_AddMenu( x, p1, p2 ) \
            (x)->lpVtbl->AddMenu( x, p1, p2 )
        #define IFileDialogCustomize_AddPushButton( x, p1, p2 ) \
            (x)->lpVtbl->AddPushButton( x, p1, p2 )
        #define IFileDialogCustomize_AddComboBox( x, p ) \
            (x)->lpVtbl->AddComboBox( x, p )
        #define IFileDialogCustomize_AddRadioButtonList( x, p ) \
            (x)->lpVtbl->AddRadioButtonList( x, p )
        #define IFileDialogCustomize_AddCheckButton( x, p1, p2, p3 ) \
            (x)->lpVtbl->AddCheckButton( x, p1, p2, p3 )
        #define IFileDialogCustomize_AddEditBox( x, p1, p2 ) \
            (x)->lpVtbl->AddEditBox( x, p1, p2 )
        #define IFileDialogCustomize_AddSeparator( x, p ) \
            (x)->lpVtbl->AddSeparator( x, p )
        #define IFileDialogCustomize_AddText( x, p1, p2 ) \
            (x)->lpVtbl->AddText( x, p1, p2 )
        #define IFileDialogCustomize_SetControlLabel( x, p1, p2 ) \
            (x)->lpVtbl->SetControlLabel( x, p1, p2 )
        #define IFileDialogCustomize_GetControlState( x, p1, p2 ) \
            (x)->lpVtbl->GetControlState( x, p1, p2 )
        #define IFileDialogCustomize_SetControlState( x, p1, p2 ) \
            (x)->lpVtbl->SetControlState( x, p1, p2 )
        #define IFileDialogCustomize_GetEditBoxText( x, p1, p2 ) \
            (x)->lpVtbl->GetEditBoxText( x, p1, p2 )
        #define IFileDialogCustomize_SetEditBoxText( x, p1, p2 ) \
            (x)->lpVtbl->SetEditBoxText( x, p1, p2 )
        #define IFileDialogCustomize_GetCheckButtonState( x, p1, p2 ) \
            (x)->lpVtbl->GetCheckButtonState( x, p1, p2 )
        #define IFileDialogCustomize_SetCheckButtonState( x, p1, p2 ) \
            (x)->lpVtbl->SetCheckButtonState( x, p1, p2 )
        #define IFileDialogCustomize_AddControlItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->AddControlItem( x, p1, p2, p3 )
        #define IFileDialogCustomize_RemoveControlItem( x, p1, p2 ) \
            (x)->lpVtbl->RemoveControlItem( x, p1, p2 )
        #define IFileDialogCustomize_RemoveAllControlItems( x, p ) \
            (x)->lpVtbl->RemoveAllControlItems( x, p )
        #define IFileDialogCustomize_GetControlItemState( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetControlItemState( x, p1, p2, p3 )
        #define IFileDialogCustomize_SetControlItemState( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetControlItemState( x, p1, p2, p3 )
        #define IFileDialogCustomize_GetSelectedControlItem( x, p1, p2 ) \
            (x)->lpVtbl->GetSelectedControlItem( x, p1, p2 )
        #define IFileDialogCustomize_SetSelectedControlItem( x, p1, p2 ) \
            (x)->lpVtbl->SetSelectedControlItem( x, p1, p2 )
        #define IFileDialogCustomize_StartVisualGroup( x, p1, p2 ) \
            (x)->lpVtbl->StartVisualGroup( x, p1, p2 )
        #define IFileDialogCustomize_EndVisualGroup( x ) \
            (x)->lpVtbl->EndVisualGroup( x )
        #define IFileDialogCustomize_MakeProminent( x, p ) \
            (x)->lpVtbl->MakeProminent( x, p )
        #define IFileDialogCustomize_SetControlItemText( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetControlItemText( x, p1, p2, p3 )
        #define IFileDialogControlEvents_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileDialogControlEvents_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileDialogControlEvents_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileDialogControlEvents_OnItemSelected( x, p1, p2, p3 ) \
            (x)->lpVtbl->OnItemSelected( x, p1, p2, p3 )
        #define IFileDialogControlEvents_OnButtonClicked( x, p1, p2 ) \
            (x)->lpVtbl->OnButtonClicked( x, p1, p2 )
        #define IFileDialogControlEvents_OnCheckButtonToggled( x, p1, p2, p3 ) \
            (x)->lpVtbl->OnCheckButtonToggled( x, p1, p2, p3 )
        #define IFileDialogControlEvents_OnControlActivating( x, p1, p2 ) \
            (x)->lpVtbl->OnControlActivating( x, p1, p2 )
        #define IFileDialog2_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IFileDialog2_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IFileDialog2_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IFileDialog2_Show( x, p ) \
            (x)->lpVtbl->Show( x, p )
        #define IFileDialog2_SetFileTypes( x, p1, p2 ) \
            (x)->lpVtbl->SetFileTypes( x, p1, p2 )
        #define IFileDialog2_SetFileTypeIndex( x, p ) \
            (x)->lpVtbl->SetFileTypeIndex( x, p )
        #define IFileDialog2_GetFileTypeIndex( x, p ) \
            (x)->lpVtbl->GetFileTypeIndex( x, p )
        #define IFileDialog2_Advise( x, p1, p2 ) \
            (x)->lpVtbl->Advise( x, p1, p2 )
        #define IFileDialog2_Unadvise( x, p ) \
            (x)->lpVtbl->Unadvise( x, p )
        #define IFileDialog2_SetOptions( x, p ) \
            (x)->lpVtbl->SetOptions( x, p )
        #define IFileDialog2_GetOptions( x, p ) \
            (x)->lpVtbl->GetOptions( x, p )
        #define IFileDialog2_SetDefaultFolder( x, p ) \
            (x)->lpVtbl->SetDefaultFolder( x, p )
        #define IFileDialog2_SetFolder( x, p ) \
            (x)->lpVtbl->SetFolder( x, p )
        #define IFileDialog2_GetFolder( x, p ) \
            (x)->lpVtbl->GetFolder( x, p )
        #define IFileDialog2_GetCurrentSelection( x, p ) \
            (x)->lpVtbl->GetCurrentSelection( x, p )
        #define IFileDialog2_SetFileName( x, p ) \
            (x)->lpVtbl->SetFileName( x, p )
        #define IFileDialog2_GetFileName( x, p ) \
            (x)->lpVtbl->GetFileName( x, p )
        #define IFileDialog2_SetTitle( x, p ) \
            (x)->lpVtbl->SetTitle( x, p )
        #define IFileDialog2_SetOkButtonLabel( x, p ) \
            (x)->lpVtbl->SetOkButtonLabel( x, p )
        #define IFileDialog2_SetFileNameLabel( x, p ) \
            (x)->lpVtbl->SetFileNameLabel( x, p )
        #define IFileDialog2_GetResult( x, p ) \
            (x)->lpVtbl->GetResult( x, p )
        #define IFileDialog2_AddPlace( x, p1, p2 ) \
            (x)->lpVtbl->AddPlace( x, p1, p2 )
        #define IFileDialog2_SetDefaultExtension( x, p ) \
            (x)->lpVtbl->SetDefaultExtension( x, p )
        #define IFileDialog2_Close( x, p ) \
            (x)->lpVtbl->Close( x, p )
        #define IFileDialog2_SetClientGuid( x, p ) \
            (x)->lpVtbl->SetClientGuid( x, p )
        #define IFileDialog2_ClearClientData( x ) \
            (x)->lpVtbl->ClearClientData( x )
        #define IFileDialog2_SetFilter( x, p ) \
            (x)->lpVtbl->SetFilter( x, p )
        #define IApplicationAssociationRegistration_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IApplicationAssociationRegistration_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IApplicationAssociationRegistration_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IApplicationAssociationRegistration_QueryCurrentDefault( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->QueryCurrentDefault( x, p1, p2, p3, p4 )
        #define IApplicationAssociationRegistration_QueryAppIsDefault( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->QueryAppIsDefault( x, p1, p2, p3, p4, p5 )
        #define IApplicationAssociationRegistration_QueryAppIsDefaultAll( x, p1, p2, p3 ) \
            (x)->lpVtbl->QueryAppIsDefaultAll( x, p1, p2, p3 )
        #define IApplicationAssociationRegistration_SetAppAsDefault( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetAppAsDefault( x, p1, p2, p3 )
        #define IApplicationAssociationRegistration_SetAppAsDefaultAll( x, p ) \
            (x)->lpVtbl->SetAppAsDefaultAll( x, p )
        #define IApplicationAssociationRegistration_ClearUserAssociations( x ) \
            (x)->lpVtbl->ClearUserAssociations( x )
        #define IApplicationAssociationRegistrationUI_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IApplicationAssociationRegistrationUI_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IApplicationAssociationRegistrationUI_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IApplicationAssociationRegistrationUI_LaunchAdvancedAssociationUI( x, p ) \
            (x)->lpVtbl->LaunchAdvancedAssociationUI( x, p )
    #endif
    #define IDelegateFolder_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IDelegateFolder_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IDelegateFolder_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IDelegateFolder_SetItemAlloc( x, p ) \
        (x)->lpVtbl->SetItemAlloc( x, p )
    #if (_WIN32_IE >= 0x0600)
        #define IBrowserFrameOptions_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IBrowserFrameOptions_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IBrowserFrameOptions_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IBrowserFrameOptions_GetFrameOptions( x, p1, p2 ) \
            (x)->lpVtbl->GetFrameOptions( x, p1, p2 )
    #endif
    #if (_WIN32_IE >= 0x0602)
        #define INewWindowManager_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define INewWindowManager_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define INewWindowManager_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define INewWindowManager_EvaluateNewWindow( x, p1, p2, p3, p4, p5, p6, p7 ) \
            (x)->lpVtbl->EvaluateNewWindow( x, p1, p2, p3, p4, p5, p6, p7 )
        #define IAttachmentExecute_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IAttachmentExecute_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IAttachmentExecute_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IAttachmentExecute_SetClientTitle( x, p ) \
            (x)->lpVtbl->SetClientTitle( x, p )
        #define IAttachmentExecute_SetClientGuid( x, p ) \
            (x)->lpVtbl->SetClientGuid( x, p )
        #define IAttachmentExecute_SetLocalPath( x, p ) \
            (x)->lpVtbl->SetLocalPath( x, p )
        #define IAttachmentExecute_SetFileName( x, p ) \
            (x)->lpVtbl->SetFileName( x, p )
        #define IAttachmentExecute_SetSource( x, p ) \
            (x)->lpVtbl->SetSource( x, p )
        #define IAttachmentExecute_SetReferrer( x, p ) \
            (x)->lpVtbl->SetReferrer( x, p )
        #define IAttachmentExecute_CheckPolicy( x ) \
            (x)->lpVtbl->CheckPolicy( x )
        #define IAttachmentExecute_Prompt( x, p1, p2, p3 ) \
            (x)->lpVtbl->Prompt( x, p1, p2, p3 )
        #define IAttachmentExecute_Save( x ) \
            (x)->lpVtbl->Save( x )
        #define IAttachmentExecute_Execute( x, p1, p2, p3 ) \
            (x)->lpVtbl->Execute( x, p1, p2, p3 )
        #define IAttachmentExecute_SaveWithUI( x, p ) \
            (x)->lpVtbl->SaveWithUI( x, p )
        #define IAttachmentExecute_ClearClientState( x ) \
            (x)->lpVtbl->ClearClientState( x )
    #endif
    #if (_WIN32_IE >= 0x0600)
        #define IShellMenuCallback_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IShellMenuCallback_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IShellMenuCallback_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IShellMenuCallback_CallbackSM( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->CallbackSM( x, p1, p2, p3, p4 )
        #define IShellMenu_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IShellMenu_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IShellMenu_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IShellMenu_Initialize( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->Initialize( x, p1, p2, p3, p4 )
        #define IShellMenu_GetMenuInfo( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->GetMenuInfo( x, p1, p2, p3, p4 )
        #define IShellMenu_SetShellFolder( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->SetShellFolder( x, p1, p2, p3, p4 )
        #define IShellMenu_GetShellFolder( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->GetShellFolder( x, p1, p2, p3, p4 )
        #define IShellMenu_SetMenu( x, p1, p2, p3 ) \
            (x)->lpVtbl->SetMenu( x, p1, p2, p3 )
        #define IShellMenu_GetMenu( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetMenu( x, p1, p2, p3 )
        #define IShellMenu_InvalidateItem( x, p1, p2 ) \
            (x)->lpVtbl->InvalidateItem( x, p1, p2 )
        #define IShellMenu_GetState( x, p ) \
            (x)->lpVtbl->GetState( x, p )
        #define IShellMenu_SetMenuToolbar( x, p1, p2 ) \
            (x)->lpVtbl->SetMenuToolbar( x, p1, p2 )
    #endif
    #define IShellRunDll_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellRunDll_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellRunDll_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellRunDll_Run( x, p ) \
        (x)->lpVtbl->Run( x, p )
    #if (NTDDI_VERSION >= 0x06000000)
        #define IKnownFolder_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IKnownFolder_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IKnownFolder_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IKnownFolder_GetId( x, p ) \
            (x)->lpVtbl->GetId( x, p )
        #define IKnownFolder_GetCategory( x, p ) \
            (x)->lpVtbl->GetCategory( x, p )
        #define IKnownFolder_GetShellItem( x, p1, p2, p3 ) \
            (x)->lpVtbl->GetShellItem( x, p1, p2, p3 )
        #define IKnownFolder_GetPath( x, p1, p2 ) \
            (x)->lpVtbl->GetPath( x, p1, p2 )
        #define IKnownFolder_SetPath( x, p1, p2 ) \
            (x)->lpVtbl->SetPath( x, p1, p2 )
        #define IKnownFolder_GetIDList( x, p1, p2 ) \
            (x)->lpVtbl->GetIDList( x, p1, p2 )
        #define IKnownFolder_GetFolderType( x, p ) \
            (x)->lpVtbl->GetFolderType( x, p )
        #define IKnownFolder_GetRedirectionCapabilities( x, p ) \
            (x)->lpVtbl->GetRedirectionCapabilities( x, p )
        #define IKnownFolder_GetFolderDefinition( x, p ) \
            (x)->lpVtbl->GetFolderDefinition( x, p )
        #define IKnownFolderManager_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IKnownFolderManager_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IKnownFolderManager_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IKnownFolderManager_FolderIdFromCsidl( x, p1, p2 ) \
            (x)->lpVtbl->FolderIdFromCsidl( x, p1, p2 )
        #define IKnownFolderManager_FolderIdToCsidl( x, p1, p2 ) \
            (x)->lpVtbl->FolderIdToCsidl( x, p1, p2 )
        #define IKnownFolderManager_GetFolderIds( x, p1, p2 ) \
            (x)->lpVtbl->GetFolderIds( x, p1, p2 )
        #define IKnownFolderManager_GetFolder( x, p1, p2 ) \
            (x)->lpVtbl->GetFolder( x, p1, p2 )
        #define IKnownFolderManager_GetFolderByName( x, p1, p2 ) \
            (x)->lpVtbl->GetFolderByName( x, p1, p2 )
        #define IKnownFolderManager_RegisterFolder( x, p1, p2 ) \
            (x)->lpVtbl->RegisterFolder( x, p1, p2 )
        #define IKnownFolderManager_UnregisterFolder( x, p ) \
            (x)->lpVtbl->UnregisterFolder( x, p )
        #define IKnownFolderManager_FindFolderFromPath( x, p1, p2, p3 ) \
            (x)->lpVtbl->FindFolderFromPath( x, p1, p2, p3 )
        #define IKnownFolderManager_FindFolderFromIDList( x, p1, p2 ) \
            (x)->lpVtbl->FindFolderFromIDList( x, p1, p2 )
        #define IKnownFolderManager_Redirect( x, p1, p2, p3, p4, p5, p6, p7 ) \
            (x)->lpVtbl->Redirect( x, p1, p2, p3, p4, p5, p6, p7 )
        #define ISharingConfigurationManager_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ISharingConfigurationManager_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ISharingConfigurationManager_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ISharingConfigurationManager_CreateShare( x, p1, p2 ) \
            (x)->lpVtbl->CreateShare( x, p1, p2 )
        #define ISharingConfigurationManager_DeleteShare( x, p ) \
            (x)->lpVtbl->DeleteShare( x, p )
        #define ISharingConfigurationManager_ShareExists( x, p ) \
            (x)->lpVtbl->ShareExists( x, p )
        #define ISharingConfigurationManager_GetSharePermissions( x, p1, p2 ) \
            (x)->lpVtbl->GetSharePermissions( x, p1, p2 )
        #define ISharingConfigurationManager_SharePrinters( x ) \
            (x)->lpVtbl->SharePrinters( x )
        #define ISharingConfigurationManager_StopSharingPrinters( x ) \
            (x)->lpVtbl->StopSharingPrinters( x )
        #define ISharingConfigurationManager_ArePrintersShared( x ) \
            (x)->lpVtbl->ArePrintersShared( x )
    #endif
    #define IPreviousVersionsInfo_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IPreviousVersionsInfo_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IPreviousVersionsInfo_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IPreviousVersionsInfo_AreSnapshotsAvailable( x, p1, p2, p3 ) \
        (x)->lpVtbl->AreSnapshotsAvailable( x, p1, p2, p3 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define IRelatedItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IRelatedItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IRelatedItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IRelatedItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IRelatedItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define IIdentityName_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IIdentityName_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IIdentityName_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IIdentityName_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IIdentityName_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define IDelegateItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDelegateItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDelegateItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDelegateItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IDelegateItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define ICurrentItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICurrentItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICurrentItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICurrentItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define ICurrentItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define ITransferMediumItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ITransferMediumItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ITransferMediumItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ITransferMediumItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define ITransferMediumItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define IUseToBrowseItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IUseToBrowseItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IUseToBrowseItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IUseToBrowseItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IUseToBrowseItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define IDisplayItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IDisplayItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IDisplayItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IDisplayItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IDisplayItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define IViewStateIdentityItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IViewStateIdentityItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IViewStateIdentityItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IViewStateIdentityItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IViewStateIdentityItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
        #define IPreviewItem_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IPreviewItem_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IPreviewItem_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IPreviewItem_GetItemIDList( x, p ) \
            (x)->lpVtbl->GetItemIDList( x, p )
        #define IPreviewItem_GetItem( x, p ) \
            (x)->lpVtbl->GetItem( x, p )
    #endif
    #define IDestinationStreamFactory_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IDestinationStreamFactory_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IDestinationStreamFactory_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IDestinationStreamFactory_GetDestinationStream( x, p ) \
        (x)->lpVtbl->GetDestinationStream( x, p )
    #define INewMenuClient_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INewMenuClient_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INewMenuClient_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INewMenuClient_IncludeItems( x, p ) \
        (x)->lpVtbl->IncludeItems( x, p )
    #define INewMenuClient_SelectAndEditItem( x, p1, p2 ) \
        (x)->lpVtbl->SelectAndEditItem( x, p1, p2 )
    #if (_WIN32_IE >= 0x0700)
        #define IInitializeWithBindCtx_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IInitializeWithBindCtx_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IInitializeWithBindCtx_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IInitializeWithBindCtx_Initialize( x, p ) \
            (x)->lpVtbl->Initialize( x, p )
        #define IShellItemFilter_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IShellItemFilter_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IShellItemFilter_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IShellItemFilter_IncludeItem( x, p ) \
            (x)->lpVtbl->IncludeItem( x, p )
        #define IShellItemFilter_GetEnumFlagsForItem( x, p1, p2 ) \
            (x)->lpVtbl->GetEnumFlagsForItem( x, p1, p2 )
    #endif
    #define INameSpaceTreeControl_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INameSpaceTreeControl_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INameSpaceTreeControl_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INameSpaceTreeControl_Initialize( x, p1, p2, p3 ) \
        (x)->lpVtbl->Initialize( x, p1, p2, p3 )
    #define INameSpaceTreeControl_TreeAdvise( x, p1, p2 ) \
        (x)->lpVtbl->TreeAdvise( x, p1, p2 )
    #define INameSpaceTreeControl_TreeUnadvise( x, p ) \
        (x)->lpVtbl->TreeUnadvise( x, p )
    #define INameSpaceTreeControl_AppendRoot( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->AppendRoot( x, p1, p2, p3, p4 )
    #define INameSpaceTreeControl_InsertRoot( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->InsertRoot( x, p1, p2, p3, p4, p5 )
    #define INameSpaceTreeControl_RemoveRoot( x, p ) \
        (x)->lpVtbl->RemoveRoot( x, p )
    #define INameSpaceTreeControl_RemoveAllRoots( x ) \
        (x)->lpVtbl->RemoveAllRoots( x )
    #define INameSpaceTreeControl_GetRootItems( x, p ) \
        (x)->lpVtbl->GetRootItems( x, p )
    #define INameSpaceTreeControl_SetItemState( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetItemState( x, p1, p2, p3 )
    #define INameSpaceTreeControl_GetItemState( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetItemState( x, p1, p2, p3 )
    #define INameSpaceTreeControl_GetSelectedItems( x, p ) \
        (x)->lpVtbl->GetSelectedItems( x, p )
    #define INameSpaceTreeControl_GetItemCustomState( x, p1, p2 ) \
        (x)->lpVtbl->GetItemCustomState( x, p1, p2 )
    #define INameSpaceTreeControl_SetItemCustomState( x, p1, p2 ) \
        (x)->lpVtbl->SetItemCustomState( x, p1, p2 )
    #define INameSpaceTreeControl_EnsureItemVisible( x, p ) \
        (x)->lpVtbl->EnsureItemVisible( x, p )
    #define INameSpaceTreeControl_SetTheme( x, p ) \
        (x)->lpVtbl->SetTheme( x, p )
    #define INameSpaceTreeControl_GetNextItem( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetNextItem( x, p1, p2, p3 )
    #define INameSpaceTreeControl_HitTest( x, p1, p2 ) \
        (x)->lpVtbl->HitTest( x, p1, p2 )
    #define INameSpaceTreeControl_GetItemRect( x, p1, p2 ) \
        (x)->lpVtbl->GetItemRect( x, p1, p2 )
    #define INameSpaceTreeControl_CollapseAll( x ) \
        (x)->lpVtbl->CollapseAll( x )
    #define INameSpaceTreeControl2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INameSpaceTreeControl2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INameSpaceTreeControl2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INameSpaceTreeControl2_Initialize( x, p1, p2, p3 ) \
        (x)->lpVtbl->Initialize( x, p1, p2, p3 )
    #define INameSpaceTreeControl2_TreeAdvise( x, p1, p2 ) \
        (x)->lpVtbl->TreeAdvise( x, p1, p2 )
    #define INameSpaceTreeControl2_TreeUnadvise( x, p ) \
        (x)->lpVtbl->TreeUnadvise( x, p )
    #define INameSpaceTreeControl2_AppendRoot( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->AppendRoot( x, p1, p2, p3, p4 )
    #define INameSpaceTreeControl2_InsertRoot( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->InsertRoot( x, p1, p2, p3, p4, p5 )
    #define INameSpaceTreeControl2_RemoveRoot( x, p ) \
        (x)->lpVtbl->RemoveRoot( x, p )
    #define INameSpaceTreeControl2_RemoveAllRoots( x ) \
        (x)->lpVtbl->RemoveAllRoots( x )
    #define INameSpaceTreeControl2_GetRootItems( x, p ) \
        (x)->lpVtbl->GetRootItems( x, p )
    #define INameSpaceTreeControl2_SetItemState( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetItemState( x, p1, p2, p3 )
    #define INameSpaceTreeControl2_GetItemState( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetItemState( x, p1, p2, p3 )
    #define INameSpaceTreeControl2_GetSelectedItems( x, p ) \
        (x)->lpVtbl->GetSelectedItems( x, p )
    #define INameSpaceTreeControl2_GetItemCustomState( x, p1, p2 ) \
        (x)->lpVtbl->GetItemCustomState( x, p1, p2 )
    #define INameSpaceTreeControl2_SetItemCustomState( x, p1, p2 ) \
        (x)->lpVtbl->SetItemCustomState( x, p1, p2 )
    #define INameSpaceTreeControl2_EnsureItemVisible( x, p ) \
        (x)->lpVtbl->EnsureItemVisible( x, p )
    #define INameSpaceTreeControl2_SetTheme( x, p ) \
        (x)->lpVtbl->SetTheme( x, p )
    #define INameSpaceTreeControl2_GetNextItem( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetNextItem( x, p1, p2, p3 )
    #define INameSpaceTreeControl2_HitTest( x, p1, p2 ) \
        (x)->lpVtbl->HitTest( x, p1, p2 )
    #define INameSpaceTreeControl2_GetItemRect( x, p1, p2 ) \
        (x)->lpVtbl->GetItemRect( x, p1, p2 )
    #define INameSpaceTreeControl2_CollapseAll( x ) \
        (x)->lpVtbl->CollapseAll( x )
    #define INameSpaceTreeControl2_SetControlStyle( x, p1, p2 ) \
        (x)->lpVtbl->SetControlStyle( x, p1, p2 )
    #define INameSpaceTreeControl2_GetControlStyle( x, p1, p2 ) \
        (x)->lpVtbl->GetControlStyle( x, p1, p2 )
    #define INameSpaceTreeControl2_SetControlStyle2( x, p1, p2 ) \
        (x)->lpVtbl->SetControlStyle2( x, p1, p2 )
    #define INameSpaceTreeControl2_GetControlStyle2( x, p1, p2 ) \
        (x)->lpVtbl->GetControlStyle2( x, p1, p2 )
    #define INameSpaceTreeControlEvents_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INameSpaceTreeControlEvents_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INameSpaceTreeControlEvents_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INameSpaceTreeControlEvents_OnItemClick( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnItemClick( x, p1, p2, p3 )
    #define INameSpaceTreeControlEvents_OnPropertyItemCommit( x, p ) \
        (x)->lpVtbl->OnPropertyItemCommit( x, p )
    #define INameSpaceTreeControlEvents_OnItemStateChanging( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnItemStateChanging( x, p1, p2, p3 )
    #define INameSpaceTreeControlEvents_OnItemStateChanged( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnItemStateChanged( x, p1, p2, p3 )
    #define INameSapceTreeControlEvents_OnSelectionChanged( x, p ) \
        (x)->lpVtbl->OnSelectionChanged( x, p )
    #define INameSpaceTreeControlEvents_OnKeyboardInput( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnKeyboardInput( x, p1, p2, p3 )
    #define INameSpaceTreeControlEvents_OnBeforeExpand( x, p ) \
        (x)->lpVtbl->OnBeforeExpand( x, p )
    #define INameSpaceTreeControlEvents_OnAfterExpand( x, p ) \
        (x)->lpVtbl->OnAfterExpand( x, p )
    #define INameSpaceTreeControlEvents_OnBeginLabelEdit( x, p ) \
        (x)->lpVtbl->OnBeginLabelEdit( x, p )
    #define INameSpaceTreeControlEvents_OnEndLabelEdit( x, p ) \
        (x)->lpVtbl->OnEndLabelEdit( x, p )
    #define INameSpaceTreeControlEvents_OnGetToolTip( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnGetToolTip( x, p1, p2, p3 )
    #define INameSpaceTreeControlEvents_OnBeforeItemDelete( x, p ) \
        (x)->lpVtbl->OnBeforeItemDelete( x, p )
    #define INameSpaceTreeControlEvents_OnItemAdded( x, p1, p2 ) \
        (x)->lpVtbl->OnItemAdded( x, p1, p2 )
    #define INameSpaceTreeControlEvents_OnItemDeleted( x, p1, p2 ) \
        (x)->lpVtbl->OnItemDeleted( x, p1, p2 )
    #define INameSpaceTreeControlEvents_OnBeforeContextMenu( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnBeforeContextMenu( x, p1, p2, p3 )
    #define INameSpaceTreeControlEvents_OnAfterContextMenu( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->OnAfterContextMenu( x, p1, p2, p3, p4 )
    #define INameSpaceTreeControlEvents_OnBeforeStateImageChange( x, p ) \
        (x)->lpVtbl->OnBeforeStateImageChange( x, p )
    #define INameSpaceTreeControlEvents_OnGetDefaultIconIndex( x, p1, p2, p3 ) \
        (x)->lpVtbl->OnGetDefaultIconIndex( x, p1, p2, p3 )
    #define INameSpaceTreeControlDropHandler_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INameSpaceTreeControlDropHandler_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INameSpaceTreeControlDropHandler_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INameSpaceTreeControlDropHandler_OnDragEnter( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->OnDragEnter( x, p1, p2, p3, p4, p5 )
    #define INameSpaceTreeControlDropHandler_OnDragOver( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->OnDragOver( x, p1, p2, p3, p4 )
    #define INameSpaceTreeControlDropHandler_OnDragPosition( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->OnDragPosition( x, p1, p2, p3, p4 )
    #define INameSpaceTreeControlDropHandler_OnDrop( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->OnDrop( x, p1, p2, p3, p4, p5 )
    #define INameSpaceTreeControlDropHandler_OnDropPosition( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->OnDropPosition( x, p1, p2, p3, p4 )
    #define INameSpaceTreeControlDropHandler_OnDragLeave( x, p ) \
        (x)->lpVtbl->OnDragLeave( x, p )
    #define INameSpaceTreeAccessible_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INameSpaceTreeAccessible_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INameSpaceTreeAccessible_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INameSpaceTreeAccessible_OnGetDefaultAccessibilityAction( x, p1, p2 ) \
        (x)->lpVtbl->OnGetDefaultAccessibilityAction( x, p1, p2 )
    #define INameSpaceTreeAccessible_OnDoDefaultAccessibilityAction( x, p ) \
        (x)->lpVtbl->OnDoDefaultAccessibilityAction( x, p )
    #define INameSpaceTreeAccessible_OnGetAccessibilityRole( x, p1, p2 ) \
        (x)->lpVtbl->OnGetAccessibilityRole( x, p1, p2 )
    #define INameSpaceTreeControlCustomDraw_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define INameSpaceTreeControlCustomDraw_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define INameSpaceTreeControlCustomDraw_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define INameSpaceTreeControlCustomDraw_PrePaint( x, p1, p2, p3 ) \
        (x)->lpVtbl->PrePaint( x, p1, p2, p3 )
    #define INameSpaceTreeControlCustomDraw_PostPaint( x, p1, p2 ) \
        (x)->lpVtbl->PostPaint( x, p1, p2 )
    #define INameSpaceTreeControlCustomDraw_ItemPrePaint( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->ItemPrePaint( x, p1, p2, p3, p4, p5, p6 )
    #define INameSpaceTreeControlCustomDraw_ItemPostPaint( x, p1, p2, p3 ) \
        (x)->lpVtbl->ItemPostPaint( x, p1, p2, p3 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define INameSpaceTreeControlFolderCapabilities_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define INameSpaceTreeControlFolderCapabilities_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define INameSpaceTreeControlFolderCapabilities_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define INameSpaceTreeControlFolderCapabilities_GetFolderCapabilities( x, p1, p2 ) \
            (x)->lpVtbl->GetFolderCapabilities( x, p1, p2 )
    #endif
    #define IPreviewHandler_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IPreviewHandler_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IPreviewHandler_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IPreviewHandler_SetWindow( x, p1, p2 ) \
        (x)->lpVtbl->SetWindow( x, p1, p2 )
    #define IPreviewHandler_SetRect( x, p ) \
        (x)->lpVtbl->SetRect( x, p )
    #define IPreviewHandler_DoPreview( x ) \
        (x)->lpVtbl->DoPreview( x )
    #define IPreviewHandler_Unload( x ) \
        (x)->lpVtbl->Unload( x )
    #define IPreviewHandler_SetFocus( x ) \
        (x)->lpVtbl->SetFocus( x )
    #define IPreviewHandler_QueryFocus( x, p ) \
        (x)->lpVtbl->QueryFocus( x, p )
    #define IPreviewHandler_TranslateAccelerator( x, p ) \
        (x)->lpVtbl->TranslateAccelerator( x, p )
    #define IPreviewHandlerFrame_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IPreviewHandlerFrame_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IPreviewHandlerFrame_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IPreviewHandlerFrame_GetWindowContext( x, p ) \
        (x)->lpVtbl->GetWindowContext( x, p )
    #define IPreviewHandlerFrame_TranslateAccelerator( x, p ) \
        (x)->lpVtbl->TranslateAccelerator( x, p )
    #if (NTDDI_VERSION >= 0x06000000)
        #define ITrayDeskBand_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ITrayDeskBand_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ITrayDeskBand_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ITrayDeskBand_ShowDeskBand( x, p ) \
            (x)->lpVtbl->ShowDeskBand( x, p )
        #define ITrayDeskBand_HideDeskBand( x, p ) \
            (x)->lpVtbl->HideDeskBand( x, p )
        #define ITrayDeskBand_IsDeskBandShown( x, p ) \
            (x)->lpVtbl->IsDeskBandShown( x, p )
        #define ITrayDeskBand_DeskBandRegistrationChanged( x ) \
            (x)->lpVtbl->DeskBandRegistrationChanged( x )
        #define IBandHost_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IBandHost_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IBandHost_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IBandHost_CreateBand( x, p1, p2, p3, p4, p5 ) \
            (x)->lpVtbl->CreateBand( x, p1, p2, p3, p4, p5 )
        #define IBandHost_SetBandAvailability( x, p1, p2 ) \
            (x)->lpVtbl->SetBandAvailability( x, p1, p2 )
        #define IBandHost_DestroyBand( x, p ) \
            (x)->lpVtbl->DestroyBand( x, p )
        #define IExplorerPaneVisibility_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IExplorerPaneVisibility_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IExplorerPaneVisibility_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IExplorerPaneVisibility_GetPaneState( x, p1, p2 ) \
            (x)->lpVtbl->GetPaneState( x, p1, p2 )
        #define IContextMenuCB_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IContextMenuCB_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IContextMenuCB_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IContextMenuCB_CallBack( x, p1, p2, p3, p4, p5, p6 ) \
            (x)->lpVtbl->CallBack( x, p1, p2, p3, p4, p5, p6 )
    #endif
    #define IDefaultExtractIconInit_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IDefaultExtractIconInit_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IDefaultExtractIconInit_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IDefaultExtractIconInit_SetFlags( x, p ) \
        (x)->lpVtbl->SetFlags( x, p )
    #define IDefaultExtractIconInit_SetKey( x, p ) \
        (x)->lpVtbl->SetKey( x, p )
    #define IDefaultExtractIconInit_SetNormalIcon( x, p1, p2 ) \
        (x)->lpVtbl->SetNormalIcon( x, p1, p2 )
    #define IDefaultExtractIconInit_SetOpenIcon( x, p1, p2 ) \
        (x)->lpVtbl->SetOpenIcon( x, p1, p2 )
    #define IDefaultExtractIconInit_SetShortcutIcon( x, p1, p2 ) \
        (x)->lpVtbl->SetShortcutIcon( x, p1, p2 )
    #define IDefaultExtractIconInit_SetDefaultIcon( x, p1, p2 ) \
        (x)->lpVtbl->SetDefaultIcon( x, p1, p2 )
    #define IExplorerCommand_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IExplorerCommand_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IExplorerCommand_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IExplorerCommand_GetTitle( x, p1, p2 ) \
        (x)->lpVtbl->GetTitle( x, p1, p2 )
    #define IExplorerCommand_GetIcon( x, p1, p2 ) \
        (x)->lpVtbl->GetIcon( x, p1, p2 )
    #define IExplorerCommand_GetToolTip( x, p1, p2 ) \
        (x)->lpVtbl->GetToolTip( x, p1, p2 )
    #define IExplorerCommand_GetCanonicalName( x, p ) \
        (x)->lpVtbl->GetCanonicalName( x, p )
    #define IExplorerCommand_GetState( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetState( x, p1, p2, p3 )
    #define IExplorerCommand_Invoke( x, p1, p2 ) \
        (x)->lpVtbl->Invoke( x, p1, p2 )
    #define IExplorerCommand_GetFlags( x, p ) \
        (x)->lpVtbl->GetFlags( x, p )
    #define IExplorerCommand_EnumSubCommands( x, p ) \
        (x)->lpVtbl->EnumSubCommands( x, p )
    #define IExplorerCommandState_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IExplorerCommandState_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IExplorerCommandState_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IExplorerCommandState_GetState( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetState( x, p1, p2, p3 )
    #define IInitializeCommand_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInitializeCommand_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInitializeCommand_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInitializeCommand_Initialize( x, p1, p2 ) \
        (x)->lpVtbl->Initialize( x, p1, p2 )
    #define IEnumExplorerCommand_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IEnumExplorerCommand_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IEnumExplorerCommand_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IEnumExplorerCommand_Next( x, p1, p2, p3 ) \
        (x)->lpVtbl->Next( x, p1, p2, p3 )
    #define IEnumExplorerCommand_Skip( x, p ) \
        (x)->lpVtbl->Skip( x, p )
    #define IEnumExplorerCommand_Reset( x ) \
        (x)->lpVtbl->Reset( x )
    #define IEnumExplorerCommand_Clone( x, p ) \
        (x)->lpVtbl->Clone( x, p )
    #define IExplorerCommandProvider_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IExplorerCommandProvider_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IExplorerCommandProvider_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IExplorerCommandProvider_GetCommands( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetCommands( x, p1, p2, p3 )
    #define IExplorerCommandProvider_GetCommand( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetCommand( x, p1, p2, p3 )
    #define IMarkupProvider_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IMarkupProvider_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IMarkupProvider_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IMarkupProvider_GetState( x, p1, p2 ) \
        (x)->lpVtbl->GetState( x, p1, p2 )
    #define IMarkupProvider_Notify( x, p1, p2, p3 ) \
        (x)->lpVtbl->Notify( x, p1, p2, p3 )
    #define IMarkupProvider_InvalidateRect( x, p1, p2 ) \
        (x)->lpVtbl->InvalidateRect( x, p1, p2 )
    #define IMarkupProvider_OnCustomDraw( x, p1, p2, p3, p4, p5, p6, p7 ) \
        (x)->lpVtbl->OnCustomDraw( x, p1, p2, p3, p4, p5, p6, p7 )
    #define IMarkupProvider_CustomDrawText( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->CustomDrawText( x, p1, p2, p3, p4, p5, p6 )
    #define IControlMarkup_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IControlMarkup_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IControlMarkup_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IControlMarkup_SetCallback( x, p ) \
        (x)->lpVtbl->SetCallback( x, p )
    #define IControlMarkup_GetCallback( x, p1, p2 ) \
        (x)->lpVtbl->GetCallback( x, p1, p2 )
    #define IControlMarkup_SetId( x, p ) \
        (x)->lpVtbl->SetId( x, p )
    #define IControlMarkup_GetId( x, p ) \
        (x)->lpVtbl->GetId( x, p )
    #define IControlMarkup_SetFonts( x, p1, p2 ) \
        (x)->lpVtbl->SetFonts( x, p1, p2 )
    #define IControlMarkup_GetFonts( x, p1, p2 ) \
        (x)->lpVtbl->GetFonts( x, p1, p2 )
    #define IControlMarkup_SetText( x, p ) \
        (x)->lpVtbl->SetText( x, p )
    #define IControlMarkup_GetText( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetText( x, p1, p2, p3 )
    #define IControlMarkup_SetLinkText( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetLinkText( x, p1, p2, p3 )
    #define IControlMarkup_GetLinkText( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->GetLinkText( x, p1, p2, p3, p4 )
    #define IControlMarkup_SetRenderFlags( x, p ) \
        (x)->lpVtbl->SetRenderFlags( x, p )
    #define IControlMarkup_GetRenderFlags( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->GetRenderFlags( x, p1, p2, p3, p4, p5 )
    #define IControlMarkup_SetThemeRenderFlags( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->SetThemeRenderFlags( x, p1, p2, p3, p4, p5 )
    #define IControlMarkup_GetState( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetState( x, p1, p2, p3 )
    #define IControlMarkup_SetState( x, p1, p2, p3 ) \
        (x)->lpVtbl->SetState( x, p1, p2, p3 )
    #define IControlMarkup_DrawText( x, p1, p2 ) \
        (x)->lpVtbl->DrawText( x, p1, p2 )
    #define IControlMarkup_SetLinkCursor( x ) \
        (x)->lpVtbl->SetLinkCursor( x )
    #define IControlMarkup_CalcIdealSize( x, p1, p2, p3 ) \
        (x)->lpVtbl->CalcIdealSize( x, p1, p2, p3 )
    #define IControlMarkup_SetFocus( x ) \
        (x)->lpVtbl->SetFocus( x )
    #define IControlMarkup_KillFocus( x ) \
        (x)->lpVtbl->KillFocus( x )
    #define IControlMarkup_IsTabbable( x ) \
        (x)->lpVtbl->IsTabbable( x )
    #define IControlMarkup_OnButtonDown( x, p ) \
        (x)->lpVtbl->OnButtonDown( x, p )
    #define IControlMarkup_OnButtonUp( x, p ) \
        (x)->lpVtbl->OnButtonUp( x, p )
    #define IControlMarkup_OnKeyDown( x, p ) \
        (x)->lpVtbl->OnKeyDown( x, p )
    #define IControlMarkup_HitTest( x, p1, p2 ) \
        (x)->lpVtbl->HitTest( x, p1, p2 )
    #define IControlMarkup_GetLinkRect( x, p1, p2 ) \
        (x)->lpVtbl->GetLinkRect( x, p1, p2 )
    #define IControlMarkup_GetControlRect( x, p ) \
        (x)->lpVtbl->GetControlRect( x, p )
    #define IControlMarkup_GetLinkCount( x, p ) \
        (x)->lpVtbl->GetLinkCount( x, p )
    #define IInitializeNetworkFolder_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInitializeNetworkFolder_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInitializeNetworkFolder_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInitializeNetworkFolder_Initialize( x, p1, p2, p3, p4, p5 ) \
        (x)->lpVtbl->Initialize( x, p1, p2, p3, p4, p5 )
    #define IOpenControlPanel_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IOpenControlPanel_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IOpenControlPanel_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IOpenControlPanel_Open( x, p1, p2, p3 ) \
        (x)->lpVtbl->Open( x, p1, p2, p3 )
    #define IOpenControlPanel_GetPath( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetPath( x, p1, p2, p3 )
    #define IOpenControlPanel_GetCurrentView( x, p ) \
        (x)->lpVtbl->GetCurrentView( x, p )
    #define ISystemCPLUpdate_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define ISystemCPLUpdate_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define ISystemCPLUpdate_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define ISystemCPLUpdate_UpdateSystemInfo( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->UpdateSystemInfo( x, p1, p2, p3, p4 )
    #define ISystemCPLUpdate_UpdateLicensingInfo( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->UpdateLicensingInfo( x, p1, p2, p3, p4 )
    #define ISystemCPLUpdate_UpdateRatingsInfo( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->UpdateRatingsInfo( x, p1, p2, p3, p4 )
    #define ISystemCPLUpdate_UpdateComputerInfo( x ) \
        (x)->lpVtbl->UpdateComputerInfo( x )
    #define IComputerInfoAdvise_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IComputerInfoAdvise_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IComputerInfoAdvise_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IComputerInfoAdvise_Advise( x, p1, p2 ) \
        (x)->lpVtbl->Advise( x, p1, p2 )
    #define IComputerInfoAdvise_Unadvise( x, p ) \
        (x)->lpVtbl->Unadvise( x, p )
    #define IComputerInfoChangeNotify_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IComputerInfoChangeNotify_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IComputerInfoChangeNotify_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IComputerInfoChangeNotify_ComputerInfoChanged( x ) \
        (x)->lpVtbl->ComputerInfoChanged( x )
    #define IFileSystemBindData_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IFileSystemBindData_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IFileSystemBindData_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IFileSystemBindData_SetFindData( x, p ) \
        (x)->lpVtbl->SetFindData( x, p )
    #define IFileSystemBindData_GetFindData( x, p ) \
        (x)->lpVtbl->GetFindData( x, p )
    #define IFileSystemBindData2_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IFileSystemBindData2_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IFileSystemBindData2_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IFileSystemBindData2_SetFindData( x, p ) \
        (x)->lpVtbl->SetFindData( x, p )
    #define IFileSystemBindData2_GetFindData( x, p ) \
        (x)->lpVtbl->GetFindData( x, p )
    #define IFileSystemBindData2_SetFileID( x, p ) \
        (x)->lpVtbl->SetFileID( x, p )
    #define IFileSystemBindData2_GetFileID( x, p ) \
        (x)->lpVtbl->GetFileID( x, p )
    #define IFileSystemBindData2_SetJunctionCLSID( x, p ) \
        (x)->lpVtbl->SetJunctionCLSID( x, p )
    #define IFileSystemBindData2_GetJunctionCLSID( x, p ) \
        (x)->lpVtbl->GetJunctionCLSID( x, p )
    #if (NTDDI_VERSION >= 0x06010000)
        #define ICustomDestinationList_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define ICustomDestinationList_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define ICustomDestinationList_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define ICustomDestinationList_SetAppID( x, p ) \
            (x)->lpVtbl->SetAppID( x, p )
        #define ICustomDestinationList_BeginList( x, p1, p2, p3 ) \
            (x)->lpVtbl->BeginList( x, p1, p2, p3 )
        #define ICustomDestinationList_AppendCategory( x, p1, p2 ) \
            (x)->lpVtbl->AppendCategory( x, p1, p2 )
        #define ICustomDestinationList_AppendKnownCategory( x, p ) \
            (x)->lpVtbl->AppendKnownCategory( x, p )
        #define ICustomDestinationList_AddUserTasks( x, p ) \
            (x)->lpVtbl->AddUserTasks( x, p )
        #define ICustomDestinationList_CommitList( x ) \
            (x)->lpVtbl->CommitList( x )
        #define ICustomDestinationList_GetRemovedDestinations( x, p1, p2 ) \
            (x)->lpVtbl->GetRemovedDestinations( x, p1, p2 )
        #define ICustomDestinationList_DeleteList( x, p ) \
            (x)->lpVtbl->DeleteList( x, p )
        #define ICustomDestinationList_AbortList( x ) \
            (x)->lpVtbl->AbortList( x )
        #define IApplicationDestinations_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IApplicationDestinations_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IApplicationDestinations_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IApplicationDestinations_SetAppID( x, p ) \
            (x)->lpVtbl->SetAppID( x, p )
        #define IApplicationDestinations_RemoveDestination( x, p ) \
            (x)->lpVtbl->RemoveDestination( x, p )
        #define IApplicationDestinations_RemoveAllDestinations( x ) \
            (x)->lpVtbl->RemoveAllDestinations( x )
        #define IApplicationDocumentLists_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IApplicationDocumentLists_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IApplicationDocumentLists_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IApplicationDocumentLists_SetAppID( x, p ) \
            (x)->lpVtbl->SetAppID( x, p )
        #define IApplicationDocumentLists_GetList( x, p1, p2, p3, p4 ) \
            (x)->lpVtbl->GetList( x, p1, p2, p3, p4 )
        #define IObjectWithAppUserModelID_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IObjectWithAppUserModelID_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IObjectWithAppUserModelID_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IObjectWithAppUserModelID_SetAppID( x, p ) \
            (x)->lpVtbl->SetAppID( x, p )
        #define IObjectWithAppUserModelID_GetAppID( x, p ) \
            (x)->lpVtbl->GetAppID( x, p )
        #define IObjectWithProgID_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IObjectWithProgID_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IObjectWithProgID_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IObjectWithProgID_SetProgID( x, p ) \
            (x)->lpVtbl->SetProgID( x, p )
        #define IObjectWithProgID_GetProgID( x, p ) \
            (x)->lpVtbl->GetProgID( x, p )
        #define IUpdateIDList_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IUpdateIDList_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IUpdateIDList_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IUpdateIDList_Update( x, p1, p2, p3 ) \
            (x)->lpVtbl->Update( x, p1, p2, p3 )
    #endif
    #define IDesktopGadget_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IDesktopGadget_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IDesktopGadget_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IDesktopGadget_RunGadget( x, p ) \
        (x)->lpVtbl->RunGadget( x, p )
    #define IHomeGroup_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IHomeGroup_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IHomeGroup_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IHomeGroup_IsMember( x, p ) \
        (x)->lpVtbl->IsMember( x, p )
    #define IHomeGroup_ShowSharingWizard( x, p1, p2 ) \
        (x)->lpVtbl->ShowSharingWizard( x, p1, p2 )
    #define IInitializeWithPropertyStore_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IInitializeWithPropertyStore_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IInitializeWithPropertyStore_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IInitializeWithPropertyStore_Initialize( x, p ) \
        (x)->lpVtbl->Initialize( x, p )
    #define IOpenSearchSource_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IOpenSearchSource_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IOpenSearchSource_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IOpenSearchSource_GetResults( x, p1, p2, p3, p4, p5, p6 ) \
        (x)->lpVtbl->GetResults( x, p1, p2, p3, p4, p5, p6 )
    #define IShellLibrary_QueryInterface( x, p1, p2 ) \
        (x)->lpVtbl->QueryInterface( x, p1, p2 )
    #define IShellLibrary_AddRef( x ) \
        (x)->lpVtbl->AddRef( x )
    #define IShellLibrary_Release( x ) \
        (x)->lpVtbl->Release( x )
    #define IShellLibrary_LoadLibraryFromItem( x, p1, p2 ) \
        (x)->lpVtbl->LoadLibraryFromItem( x, p1, p2 )
    #define IShellLibrary_LoadLibraryFromKnownFolder( x, p1, p2 ) \
        (x)->lpVtbl->LoadLibraryFromKnownFolder( x, p1, p2 )
    #define IShellLibrary_AddFolder( x, p ) \
        (x)->lpVtbl->AddFolder( x, p )
    #define IShellLibrary_RemoveFolder( x, p ) \
        (x)->lpVtbl->RemoveFolder( x, p )
    #define IShellLibrary_GetFolders( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetFolders( x, p1, p2, p3 )
    #define IShellLibrary_ResolveFolder( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->ResolveFolder( x, p1, p2, p3, p4 )
    #define IShellLibrary_GetDefaultSaveFolder( x, p1, p2, p3 ) \
        (x)->lpVtbl->GetDefaultSaveFolder( x, p1, p2, p3 )
    #define IShellLibrary_SetDefaultSaveFolder( x, p1, p2 ) \
        (x)->lpVtbl->SetDefaultSaveFolder( x, p1, p2 )
    #define IShellLibrary_GetOptions( x, p ) \
        (x)->lpVtbl->GetOptions( x, p )
    #define IShellLibrary_SetOptions( x, p1, p2 ) \
        (x)->lpVtbl->SetOptions( x, p1, p2 )
    #define IShellLibrary_GetFolderType( x, p ) \
        (x)->lpVtbl->GetFolderType( x, p )
    #define IShellLibrary_SetFolderType( x, p ) \
        (x)->lpVtbl->SetFolderType( x, p )
    #define IShellLibrary_GetIcon( x, p ) \
        (x)->lpVtbl->GetIcon( x, p )
    #define IShellLibrary_SetIcon( x, p ) \
        (x)->lpVtbl->SetIcon( x, p )
    #define IShellLibrary_Commit( x ) \
        (x)->lpVtbl->Commit( x )
    #define IShellLibrary_Save( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->Save( x, p1, p2, p3, p4 )
    #define IShellLibrary_SaveInKnownFolder( x, p1, p2, p3, p4 ) \
        (x)->lpVtbl->SaveInKnownFolder( x, p1, p2, p3, p4 )
    #if (NTDDI_VERSION >= 0x06000000)
        #define IAssocHandlerInvoker_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IAssocHandlerInvoker_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IAssocHandlerInvoker_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IAssocHandlerInvoker_SupportsSelection( x ) \
            (x)->lpVtbl->SupportsSelection( x )
        #define IAssocHandlerInvoker_Invoke( x ) \
            (x)->lpVtbl->Invoke( x )
        #define IAssocHandler_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IAssocHandler_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IAssocHandler_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IAssocHandler_GetName( x, p ) \
            (x)->lpVtbl->GetName( x, p )
        #define IAssocHandler_GetUIName( x, p ) \
            (x)->lpVtbl->GetUIName( x, p )
        #define IAssocHandler_GetIconLocation( x, p1, p2 ) \
            (x)->lpVtbl->GetIconLocation( x, p1, p2 )
        #define IAssocHandler_IsRecommended( x ) \
            (x)->lpVtbl->IsRecommended( x )
        #define IAssocHandler_MakeDefault( x, p ) \
            (x)->lpVtbl->MakeDefault( x, p )
        #define IAssocHandler_Invoke( x, p ) \
            (x)->lpVtbl->Invoke( x, p )
        #define IAssocHandler_CreateInvoker( x, p1, p2 ) \
            (x)->lpVtbl->CreateInvoker( x, p1, p2 )
        #define IEnumAssocHandlers_QueryInterface( x, p1, p2 ) \
            (x)->lpVtbl->QueryInterface( x, p1, p2 )
        #define IEnumAssocHandlers_AddRef( x ) \
            (x)->lpVtbl->AddRef( x )
        #define IEnumAssocHandlers_Release( x ) \
            (x)->lpVtbl->Release( x )
        #define IEnumAssocHandlers_Next( x, p1, p2, p3 ) \
            (x)->lpVtbl->Next( x, p1, p2, p3 )
    #endif
#endif

/* Functions in SHELL32.DLL */
SHSTDAPI    SHCreateDefaultExtractIcon( REFIID, void ** );
SHSTDAPI_( PIDLIST_ABSOLUTE )   SHSimpleIDListFromPath( LPCWSTR );
#if (_WIN32_IE >= 0x0700)
SHSTDAPI    SHAddDefaultPropertiesByExt( PCWSTR, IPropertyStore * );
SHSTDAPI    SHCreateDefaultPropertiesOp( IShellItem *, IFileOperation ** );
SHSTDAPI    SHCreateItemFromIDList( PCIDLIST_ABSOLUTE, REFIID, void ** );
SHSTDAPI    SHCreateItemFromParsingName( PCWSTR, IBindCtx *, REFIID, void ** );
SHSTDAPI    SHCreateItemFromRelativeName( IShellItem *, PCWSTR, IBindCtx *, REFIID, void ** );
SHSTDAPI    SHCreateItemWithParent( PCIDLIST_ABSOLUTE, IShellFolder *, PCUITEMID_CHILD, REFIID, void ** );
SHSTDAPI    SHCreateShellItemArray( PCIDLIST_ABSOLUTE, IShellFolder *, UINT, PCUITEMID_CHILD_ARRAY, IShellItemArray ** );
SHSTDAPI    SHCreateShellItemArrayFromDataObject( IDataObject *, REFIID, void ** );
SHSTDAPI    SHCreateShellItemArrayFromIDList( UINT, PCIDLIST_ABSOLUTE_ARRAY, IShellItemArray ** );
SHSTDAPI    SHCreateShellItemArrayFromShellItem( IShellItem *, REFIID, void ** );
SHSTDAPI    SHSetDefaultProperties( HWND, IShellItem *, DWORD, IFileOperationProgressSink * );
#endif
#if (NTDDI_VERSION >= 0x06000000)
SHSTDAPI    SHAssocEnumHandlers( LPCWSTR, ASSOC_FILTER, IEnumAssocHandlers ** );
SHSTDAPI    SHCreateAssociationRegistration( REFIID, void ** );
SHSTDAPI    SHCreateItemInKnownFolder( REFKNOWNFOLDERID, DWORD, PCWSTR, REFIID, void ** );
SHSTDAPI    SHGetIDListFromObject( IUnknown *, PIDLIST_ABSOLUTE * );
SHSTDAPI    SHGetItemFromObject( IUnknown *, REFIID, void ** );
SHSTDAPI    SHGetNameFromIDList( PCIDLIST_ABSOLUTE, SIGDN, PWSTR * );
SHSTDAPI    SHGetPropertyStoreFromIDList( PCIDLIST_ABSOLUTE, GETPROPERTYSTOREFLAGS, REFIID, void ** );
SHSTDAPI    SHGetPropertyStoreFromParsingName( PCWSTR, IBindCtx *, GETPROPERTYSTOREFLAGS, REFIID, void ** );
SHSTDAPI    SHGetTemporaryPropertyForItem( IShellItem *, REFPROPERTYKEY, PROPVARIANT * );
SHSTDAPI    SHSetTemporaryPropertyForItem( IShellItem *, REFPROPERTYKEY, REFPROPVARIANT );
#endif
#if (NTDDI_VERSION >= 0x06010000)
SHSTDAPI    GetCurrentProcessExplicitAppUserModelID( PWSTR * );
SHSTDAPI    SHAssocEnumHandlersForProtocolByApplication( PCWSTR, REFIID, void ** );
SHSTDAPI    SHGetItemFromDataObject( IDataObject *, DATAOBJ_GET_ITEM_FLAGS, REFIID, void ** );
SHSTDAPI    SetCurrentProcessExplicitAppUserModelID( PCWSTR );
#endif
#if (NTDDI_VERSION >= 0x06010000) && (_WIN32_IE >= 0x0700)
SHSTDAPI    SHShowManageLibraryUI( IShellItem *, HWND, LPCWSTR, LPCWSTR, LIBRARYMANAGEDIALOGOPTIONS );
SHSTDAPI    SHResolveLibrary( IShellItem * );
#endif

#if (NTDDI_VERSION >= 0x06000000)

__inline void FreeKnownFolderDefinitionFields( KNOWNFOLDER_DEFINITION *x )
{
    CoTaskMemFree( x->pszName );
    CoTaskMemFree( x->pszDescription );
    CoTaskMemFree( x->pszRelativePath );
    CoTaskMemFree( x->pszParsingName );
    CoTaskMemFree( x->pszTooltip );
    CoTaskMemFree( x->pszLocalizedName );
    CoTaskMemFree( x->pszIcon );
    CoTaskMemFree( x->pszSecurity );
}

#endif /* (NTDDI_VERSION >= 0x06000000) */

:include cplusepi.sp

#endif /* __shobjidl_h__ */