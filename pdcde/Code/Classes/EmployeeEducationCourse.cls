VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeEducationCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lnCourseid As Long
Private strEmpCode As String
Private strcode As String
Private strcourse As String
Private dtCfrom As Date
Private dtTo As Date
Private strlevel As String
Private straward As String
Private strcomments As String
Private lnemployeeid As Long
Private strinstitution As String


Public Property Get courseid() As Long
courseid = lnCourseid
End Property
Public Property Let courseid(vnew As Long)
lnCourseid = vnew
End Property

Public Property Get EmpCode() As String
EmpCode = strEmpCode
End Property

Public Property Let EmpCode(vnew As String)
strEmpCode = vnew
End Property

Public Property Get Code() As String
Code = strcode
End Property

Public Property Let Code(vnew As String)
strcode = vnew
End Property

Public Property Get Course() As String
Course = strcourse
End Property

Public Property Let Course(vnew As String)
strcourse = vnew
End Property

Public Property Get cFrom() As Date
cFrom = dtCfrom
End Property

Public Property Let cFrom(vnew As Date)
dtCfrom = vnew
End Property

Public Property Get cTo() As Date
cTo = dtTo
End Property
Public Property Let cTo(vnew As Date)
dtTo = vnew
End Property

Public Property Get Level() As String
Level = strlevel
End Property

Public Property Let Level(vnew As String)
strlevel = vnew
End Property

Public Property Get Award() As String
Award = straward
End Property
Public Property Let Award(vnew As String)
straward = vnew
End Property

Public Property Get Comments() As String
Comments = strcomments
End Property
Public Property Let Comments(vnew As String)
strcomments = vnew
End Property

Public Property Get EmployeeID() As Long
EmployeeID = lnemployeeid
End Property

Public Property Let EmployeeID(vnew As Long)
lnemployeeid = vnew
End Property

Public Property Get Institution() As String
Institution = strinstitution
End Property

Public Property Let Institution(vnew As String)
strinstitution = vnew
End Property
