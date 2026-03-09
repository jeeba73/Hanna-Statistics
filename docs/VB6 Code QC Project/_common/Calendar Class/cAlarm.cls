VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlarm"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eDateScrollMode
    dsDate = 0
    dsTime
    dsYear
    dsMonth
    dsMonthYear
End Enum

Public Enum eAlarmType
    ccSystem = -1
    ccPopup = 0
    ccBeep
    ccSoundFile
    ccExecute
End Enum

Public Enum eUpdate
    ccNever = 0
    ccDaily
    ccWeekly
    ccMonthly
End Enum

Public UID As Double
Public dateTime As Date
Public alarmTime As Date
Public alarmType As eAlarmType
Public filename As String
Public repeat As Integer
Public message As String
Public presetMins As Integer

