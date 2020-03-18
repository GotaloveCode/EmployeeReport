VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EducationHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lngEducationID As Long
Private streducationCode As String
Private streducationName As String
Private lnEployeeid As Long
Private strEducationlevel As String
Private strinstitution As String
Private dtFrom As Date
Private dtTo As Date
Private lnawardid As Long
Private straward As String
Private strcomments As String
Private Myeducation As EducationCourse
Private Myeducations As EducationCourses

Public Property Let Educationid(ByVal vnew As Long)
    lngEducationID = vnew
End Property

Public Property Get Educationid() As Long
    Educationid = lngEducationID
End Property

Public Property Let EducationCode(ByVal vnew As String)
    streducationCode = vnew
End Property

Public Property Get EducationCode() As String
    EducationCode = streducationCode
End Property

Public Property Let educationName(ByVal vnew As String)
    streducationName = vnew
End Property

Public Property Get educationName() As String
    educationName = streducationName
End Property
''
Public Property Let EmployeeID(ByVal vnew As Long)
    lnEployeeid = vnew
End Property

Public Property Get EmployeeID() As Long
    EmployeeID = lnEployeeid
End Property

Public Property Let educationlevel(ByVal vnew As String)
    strEducationlevel = vnew
End Property
Public Property Get educationlevel() As String
    educationlevel = strEducationlevel
End Property

Public Property Get Institution() As String
    Institution = strinstitution
End Property

Public Property Let Institution(ByVal vnew As String)
    strinstitution = vnew
End Property

Public Property Get From() As Date
    From = dtFrom
End Property

Public Property Let From(ByVal vnew As Date)
    dtFrom = vnew
End Property

Public Property Get dto() As Date
    dto = dtTo
End Property

Public Property Let dto(ByVal vnew As Date)
    dtTo = vnew
End Property

Public Property Get Award() As String
    Award = straward
End Property

Public Property Let Award(ByVal vnew As String)
    straward = vnew
End Property

Public Property Get awardid() As Long
    awardid = lnawardid
End Property

Public Property Let awardid(ByVal vnew As Long)
    lnawardid = vnew
End Property

Public Property Get Comments() As String
    Comments = strcomments
End Property

Public Property Let Comments(ByVal vnew As String)
    strcomments = vnew
End Property



