VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbSaQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ----------------------------------------------------------------------------
' Workbook SaQ.xlsb(wbSaQ): Stk (mStck) and Queue (mQue) services.
'
' Standard Modules
' - mQue    Queue services
' - mStck   Stk services
' - mTest   Regression-Test
'
' W. Rauschenberger Berlin Jan 2013
' ----------------------------------------------------------------------------
Private Const HOSTED_RAWS = "mStk,mQ,clsStk, clsQ"

Private Sub Workbook_Open()
    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Serviced-Folder' path and either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook.
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '~~ The following statement invokes the named service, provided the Workbook
    '~~ has been opened from within the 'Serviced-Folder' path and either the
    '~~ CompMan development instance or the CompMan Addin instance is available
    '~~ as the servicing Workbook.
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub

