VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Viewer"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8100
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Demonstrated by: Julius Enerio
' Requirements: Make sure the Adobe Acrobat 7.0 Control Type Library (AcroPDF.dll) is visible in your toolbox
' Right Click on the toolbox then select Components. Select Adobe Acrobat 7.0 Control Type Library then click OK
' No need to add the control in the form though. Just use the code.
' Read More on how to use the browser control by reading the Interapplication communication API Reference found by searching in on Google

Option Explicit 'Which means all variables must be declared before it can be used in the program

Private m_objPDF As AcroPDFLibCtl.AcroPDF 'Declare an object of type AcroPDF
Private m_strFilePath As String 'Declare a string for the PDF Filename and Path

'On Form Load...
Private Sub Form_Load()
   m_strFilePath = App.Path & "\01 SQL Server Tour.pdf" 'Change this to the path and filename of your PDF File
   Set m_objPDF = Controls.Add("AcroPDF.PDF.1", "Test") 'This will add the PDF Browser control to the form on runtime. The "Test" is the control's name
   Set m_objPDF.Container = Frame1 'Attach the PDF Browser control to a container.
   'A Container can be a Frame, PictureBox, or SSTab Control. In this code, I used a Frame.
   
End Sub

'On Form Activate
Private Sub Form_Activate()
   'Load the PDF file specified in m_strFilePath.
   'Make sure to do this before doing any changes to the browser controls view/layout
   m_objPDF.LoadFile m_strFilePath
   
   'Set whether a toolbar will appear in the viewer. True to show, False to Hide.
   m_objPDF.setShowToolbar False
   
   'Sets the Layout Mode for a page view according to the specified string.
   'DontCare — use the current user preference
   'SinglePage — use single page mode (as it would have appeared in pre-Acrobat 3.0 viewers)
   'OneColumn — use one-column continuous mode
   'TwoColumnLeft — use two-column continuous mode with the first page on the left
   'TwoColumnRight — use two-column continuous mode with the first page on the right
   m_objPDF.setLayoutMode "SinglePage"
   
   'Sets the page mode in which a document is to be opened
   'PDDontCare: 0 — leave the view mode as it is
   'PDUseNone: 1 — display without bookmarks or thumbnails
   'PDUseThumbs: 2 — display using thumbnails
   'PDUseBookmarks: 3 — display using bookmarks
   m_objPDF.setPageMode "none"
   
   'Set the Zoom view according to the value specified. ranges from 0 and onwards
   m_objPDF.setZoom 75
   
   'Move and Resize the object in relation to its container/form
   With m_objPDF
      .Move 125, 175, 7800, 8415 'x-position, y-position, width, height
   End With
   
   'Show the Browser Control
   m_objPDF.Visible = True
End Sub

'On Form Unload
Private Sub Form_Unload(Cancel As Integer)
 
 'The Browser control will load a blank
 m_objPDF.LoadFile ""
 
 'Set object to nothing
 Set m_objPDF = Nothing
End Sub
