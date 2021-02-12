VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStrList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Constants
Const INITIAL_SIZE = 10     ' Initial size of the array

' Class variables
Private mElements As Integer
Private mStrArray() As String


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Initialization and Destruction Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Allocate initial amount of memory to the array
Private Sub Class_Initialize()
    ReDim Preserve mStrArray(INITIAL_SIZE)
End Sub

' Deallocate the entire array
Private Sub Class_Terminate()
    Clear
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Properties
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Return the current array size
Public Property Get Size() As Long
    Size = UBound(mStrArray)
End Property

' Return current amount of elements in the array
Public Property Get Count() As Long
    Count = mElements
End Property

' Return the value found at the specified location
Public Property Get Item(index As Integer) As String
    If index < Size() Then
        Item = mStrArray(index)
    End If
End Property

' Assign a new value to the specified item
Public Property Let Item(index As Integer, newValue As String)
    If index < Size() Then
        mStrArray(index) = newValue
        If index >= mElements Then mElements = index + 1
    End If
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Add a new item to the array
Public Sub Add(newValue As String)
    ' Check the size of the array
    Resize
    
    ' Add the value at the next index
    mStrArray(mElements) = newValue
    mElements = mElements + 1
End Sub

' Add a new item to the array only if it doesn't already exist
Public Sub AddNew(newValue As String)
    If Find(newValue) >= 0 Then
        Exit Sub
    Else
        Add newValue
    End If
End Sub

' Resize the array to fit more items
Private Sub Resize()
    Dim arraySize As Long
    arraySize = Size()
    If Count() = arraySize Then
        ReDim Preserve mStrArray(arraySize + INITIAL_SIZE)
    End If
End Sub

' Erase the array
Public Sub Clear()
    Erase mStrArray
    mElements = 0
End Sub

' Find the specified item in the list;
' return its index if found, -1 otherwise
Public Function Find(strItem As String) As Integer
    ' Find the specified item
    Dim i As Integer
    For i = 0 To Count - 1
        If Item(i) = strItem Then
            Find = i
            Exit Function
        End If
    Next
    
    ' No match was found -- return -1
    Find = -1
End Function

' Display contents of the list on the screen
Public Sub Display(Optional Title As String = "")
    Dim i As Integer, str As String
    For i = 0 To Count - 1
        str = str & Item(i) & vbCrLf
    Next
    MsgBox str, vbInformation, Title
End Sub

