Attribute VB_Name = "Module2"
'------------------------------------------------------------------------------
' WINPERF.H / PDH.DLL constants and defines
' Matt Pietrek - March 1998 MSJ
'------------------------------------------------------------------------------

Enum PERF_DETAIL
PERF_DETAIL_NOVICE = 100      ' The uninformed can understand it
PERF_DETAIL_ADVANCED = 200    ' For the advanced user
PERF_DETAIL_EXPERT = 300      ' For the expert user
PERF_DETAIL_WIZARD = 400      ' For the system designer
End Enum

Enum PDH_STATUS
PDH_CSTATUS_VALID_DATA = &H0
PDH_CSTATUS_NEW_DATA = &H1
PDH_CSTATUS_NO_MACHINE = &H800007D0
PDH_CSTATUS_NO_INSTANCE = &H800007D1
PDH_MORE_DATA = &H800007D2
PDH_CSTATUS_ITEM_NOT_VALIDATED = &H800007D3
PDH_RETRY = &H800007D4
PDH_NO_DATA = &H800007D5
PDH_CALC_NEGATIVE_DENOMINATOR = &H800007D6
PDH_CALC_NEGATIVE_TIMEBASE = &H800007D7
PDH_CALC_NEGATIVE_VALUE = &H800007D8
PDH_DIALOG_CANCELLED = &H800007D9
PDH_CSTATUS_NO_OBJECT = &HC0000BB8
PDH_CSTATUS_NO_COUNTER = &HC0000BB9
PDH_CSTATUS_INVALID_DATA = &HC0000BBA
PDH_MEMORY_ALLOCATION_FAILURE = &HC0000BBB
PDH_INVALID_HANDLE = &HC0000BBC
PDH_INVALID_ARGUMENT = &HC0000BBD
PDH_FUNCTION_NOT_FOUND = &HC0000BBE
PDH_CSTATUS_NO_COUNTERNAME = &HC0000BBF
PDH_CSTATUS_BAD_COUNTERNAME = &HC0000BC0
PDH_INVALID_BUFFER = &HC0000BC1
PDH_INSUFFICIENT_BUFFER = &HC0000BC2
PDH_CANNOT_CONNECT_MACHINE = &HC0000BC3
PDH_INVALID_PATH = &HC0000BC4
PDH_INVALID_INSTANCE = &HC0000BC5
PDH_INVALID_DATA = &HC0000BC6
PDH_NO_DIALOG_DATA = &HC0000BC7
PDH_CANNOT_READ_NAME_STRINGS = &HC0000BC8
End Enum

Global Const ERROR_SUCCESS = 0

Declare Function PdhVbGetOneCounterPath _
    Lib "PDH.DLL" _
    (ByVal PathString As String, _
    ByVal PathLength As Long, _
    ByVal DetailLevel As Long, _
    ByVal CaptionString As String) _
    As Long
    
Declare Function PdhVbCreateCounterPathList _
        Lib "PDH.DLL" _
        (ByVal PERF_DETAIL As Long, _
         ByVal CaptionString As String) _
        As Long

Declare Function PdhVbGetCounterPathFromList _
        Lib "PDH.DLL" _
        (ByVal Index As Long, _
         ByVal Buffer As String, _
         ByVal BufferLength As Long) _
        As Long

Declare Function PdhOpenQuery _
    Lib "PDH.DLL" _
    (ByVal Reserved As Long, _
    ByVal dwUserData As Long, _
    ByRef hQuery As Long) _
    As PDH_STATUS

Declare Function PdhCloseQuery _
    Lib "PDH.DLL" _
    (ByVal hQuery As Long) _
    As PDH_STATUS

Declare Function PdhVbAddCounter _
    Lib "PDH.DLL" _
    (ByVal QueryHandle As Long, _
    ByVal CounterPath As String, _
    ByRef CounterHandle As Long) _
    As PDH_STATUS

Declare Function PdhCollectQueryData _
    Lib "PDH.DLL" _
    (ByVal QueryHandle As Long) _
    As PDH_STATUS
    
Declare Function PdhVbIsGoodStatus _
    Lib "PDH.DLL" _
    (ByVal StatusValue As Long) _
    As Long
    
Declare Function PdhVbGetDoubleCounterValue _
    Lib "PDH.DLL" _
    (ByVal CounterHandle As Long, _
    ByRef CounterStatus As Long) _
    As Double
    


