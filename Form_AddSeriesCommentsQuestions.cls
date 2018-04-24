VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_AddSeriesCommentsQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**
'* @class Form_AddSeriesCommentsQuestions
'* This is a class module to define the behaviors for the AddSeriesCommentsQuestions form,
'* which connects a convenient user interface form to the database records from the
'* series_commentsquestions table, allowing the user to review posted comments/questions
'* and to create, edit and post new comments/questions on a given Series record.
'*
'* @author Charles Johnson
'* @version 2017.0721
'**

'Option directives determine how the Visual Basic for Applications interpreter will run
'Option Compare Database instructs the interpreter to use the locale of the current db
'to perform string comparisons
'Option Explicit requires all variables used in a procedure to be declared with explicit
'Dim or ReDim statements
Option Compare Database
Option Explicit

Private initCollectionNumber As String
Private initCommentType As String
Private initCommentText As String

'**
'* cmdPost_Click: When the user clicks the cmdPost ("Post Comment") button...
'* (1) Saves the current record (thus "posting" a new comment)
'* (2) Triggers a method to visibly update comment count back on the AddASeries Form
'* (3) Closes the current instance of the AddSeriesCommentsQuestions form, returning AddASeries form
'*
'* @uses Form_AddASeries.ResetCollectionNumber
'**
Private Sub cmdPost_Click()
    DoCmd.RunCommand acCmdSaveRecord
    'Call Forms("AddASeries").ResetCollectionNumber
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

'**
'* Form_Open: When the user opens a new instance of the AddSeriesCommentsQuestions form...
'* ... Parse the OpenArgs string, if any, provided, which calling modules like Form_AddASeries
'* can use to provide initial values for CollectionNumber, CommentTyext, CommentType, etc. if
'* desired, to be pre-filled for the user.
'**
Private Sub Form_Open(Cancel As Integer)
    Dim aParts As Variant
    Dim sPart As Variant
    Dim aKeyVal As Variant
    
    If Not IsNull(Me.OpenArgs) Then
        aParts = Split(Me.OpenArgs, "&")
        For Each sPart In aParts
            aKeyVal = Split(sPart, "=", 2)
            If (UBound(aKeyVal) < 1) Then
                initCollectionNumber = aKeyVal(0)
            ElseIf aKeyVal(0) = "CollectionNumber" Then
                initCollectionNumber = aKeyVal(1)
            ElseIf aKeyVal(0) = "CommentText" Then
                initCommentText = aKeyVal(1)
            ElseIf aKeyVal(0) = "CommentType" Then
                initCommentType = aKeyVal(1)
            End If
        Next sPart
    End If
End Sub

'**
'* Form_Current: When the user navigates or is sent to a given Current record on the form...
'* ... Set up the edit fields for the Current record. If this is a *NEW* comment/question to
'* be posted, check whether any pre-filled field values were provided by OpenArgs, and fill
'* in those fields with initial values, if so; fill in a signature for CommentAuthor using
'* getUserName; then open all fields for editing. If this is an *EXISTING* comment/question
'* being reviewed, then don't alter the contents of any fields, and lock them as read-only.
'*
'* @uses getUserName
'* @uses Update_txtSeriesTitle_from_CollectionNumber
'**
Private Sub Form_Current()
    'Is this a new Comment/Question? If so, were we provided any initial data to fill in?
    If Me.NewRecord Then
        If (Len(initCollectionNumber) > 0) Then
            Me!CollectionNumber.value = initCollectionNumber
            initCollectionNumber = ""
            
            Me!CommentAuthor.value = getUserName
        End If
        
        If (Len(initCommentType) > 0) Then
            Me!CommentType.value = initCommentType
            initCommentType = ""
        End If
        
        If (Len(initCommentText) > 0) Then
            Me!CommentText.value = initCommentText
            initCommentText = ""
        End If
    End If
    
    'Update the txtSeriesTitle utility read-out and editing box
    'If Not IsNull(Me!CollectionNumber.Value) Then
        'Update_txtSeriesTitle_from_CollectionNumber Me!ACCN.Value, Me!Agency.Value
    'Else
        'Update_txtSeriesTitle_from_CollectionNumber "", ""
    'End If
    
    'This form allows you to REVIEW existing comments/questions, and to EDIT/POST new ones
    'If you need to EDIT an existing comment/question, directly edit the DB
    Me!ACCN.Enabled = Me.NewRecord
    Me!Agency.Enabled = Me.NewRecord
    Me!CommentAuthor.Enabled = Me.NewRecord
    Me!cmbCommentType.Enabled = Me.NewRecord
    Me!CommentText.Enabled = Me.NewRecord
    Me!CommentDate.Enabled = Me.NewRecord
    
End Sub

'**
'* cmdRecordEdit_Click: When the user clicks the cmdRecordEdit ("Save") button in the txtSeriesTitle text field...
'* Executes a SQL query to edit the SeriesTitle of the grcallnumbers record that this Comment/Question references
'* to the new value that the user provided in the txtSeriesTitle text field. This allows for quick, convenient editing
'* of (for example) typos in SeriesTitle entries, while adding a comment documenting the change that you made.
'**
Private Sub cmdRecordEdit_Click()
    
    'If Not IsNull(txtSeriesTitle.Value) Then
    '    CurrentDb.Execute "UPDATE grcallnumbers SET SeriesTitle = '" & Replace(txtSeriesTitle.Value, "'", "''") & "' WHERE CollectionNumber = '" & Replace(Me!CollectionNumber.Value, "'", "''") & "'"
    '    MsgBox "Record updated in grcallnumbers!"
    'End If
End Sub

'**
'* CollectionNumber_KeyUp: When the user finishes pressing a key in the CollectionNumber text edit field...
'* ... Update the contents of the txtSeriesTitle utility read-out according to the up-to-date (.Text) value
'* of the CollectionNumber field (so the read-out shows the titles matching the current contents of that field)
'*
'* @uses Update_txtSeriesTitle_from_Collection_Number
'**
Private Sub CollectionNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    
    'Update_txtSeriesTitle_from_CollectionNumber Me!CollectionNumber.Text
        
End Sub

'**
'* Update_txtSeriesTitle_from_CollectionNumber: Update the contents of the txtSeriesTitle utility read-out / edit box
'* using the provided text CollectionNumber text value (for example, "GR-SOS-1")
'*
'* @param String Text
'**
Private Sub Update_txtSeriesTitle_from_CollectionNumber(Text As String)

    Dim oRS As DAO.Recordset
    'Set oRs = CurrentDb.OpenRecordset("SELECT * FROM grcallnumbers WHERE CollectionNumber LIKE '" & Replace(Text, "'", "''") & "*'")
    
    'txtSeriesTitle.Value = ""
    'Do Until oRs.EOF
        'txtSeriesTitle.Value = txtSeriesTitle.Value & oRs!SeriesTitle & Chr$(13) & Chr$(10)
        'oRs.MoveNext
    'Loop

    Set oRS = Nothing
End Sub
