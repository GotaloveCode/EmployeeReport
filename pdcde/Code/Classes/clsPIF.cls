VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lngid As Long
Private strCategory As String
Public Property Get ID() As Long
    ID = lngid
End Property
    
Public Property Let ID(ByVal vNew As Long)
    lngid = vNew
End Property

Public Property Get Category() As String
    Category = strCategory
End Property
    
Public Property Let Category(ByVal vNew As String)
    strCategory = vNew
End Property
