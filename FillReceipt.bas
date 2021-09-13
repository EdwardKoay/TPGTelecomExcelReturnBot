Attribute VB_Name = "Module37"
Option Explicit

Sub FillReceipt()
Attribute FillReceipt.VB_ProcData.VB_Invoke_Func = "Q\n14"

    Dim LastRow As String
    Dim R As Variant
    Dim C As Variant
    Dim ISP As String
    Dim CID As String
    Dim UN As String
    Dim Reason As String
    
    'Dim startTime As Double
    'startTime = Timer
    
    Dim wApp As Word.Application
    Dim wDoc As Word.document
    Dim rngStory As Word.Range
    Dim lngJunk As Long
    
    LastRow = Range("A1048576").End(xlUp).Row()
    Cells(LastRow, "A").Select
    
    R = ActiveCell.Row
    C = ActiveCell.Column
    
    CID = Cells(R, C + 2).Value
    UN = Cells(R, C + 3).Value
    ISP = Cells(R, C + 9).Value
    Reason = Cells(R, C + 8).Value
    
    'Create a new instance of Word & make it visible
    Set wApp = CreateObject("Word.Application")
    wApp.Visible = True
    
    'Checks ISP then opens & fills the relevant document
    If (ISP = "TPG") Then
    
        Set wDoc = wApp.Documents.Open("Z:\03_ADSL\Modreq\Ed\Returns\TPG Receipt")
        
         lngJunk = wDoc.Sections(1).headers(1).Range.StoryType
    
          'Iterate through all story types in the current document
        
          For Each rngStory In wDoc.StoryRanges
        
            'Iterate through all linked stories
        
            Do
        
              With rngStory.Find
        
                .Text = "<<CID>>"
        
                .Replacement.Text = CID
                
                .Execute Replace:=wdReplaceAll
                
                .Text = "<<UN>>"
        
                .Replacement.Text = UN
                
                .Execute Replace:=wdReplaceAll
                
                .Text = "<<REASON>>"
        
                .Replacement.Text = Reason
                
                .Execute Replace:=wdReplaceAll
        
              End With
              
              'Get next linked story (if any)
        
              Set rngStory = rngStory.NextStoryRange
        
            Loop Until rngStory Is Nothing
            Next
    Else
    
        Set wDoc = wApp.Documents.Open("Z:\03_ADSL\Modreq\Ed\Returns\iiNet Receipt")
        
         lngJunk = wDoc.Sections(1).headers(1).Range.StoryType
    
          'Iterate through all story types in the current document
        
          For Each rngStory In wDoc.StoryRanges
        
            'Iterate through all linked stories
        
            Do
        
              With rngStory.Find
        
                .Text = "<<CID>>"
        
                .Replacement.Text = CID
                
                .Execute Replace:=wdReplaceAll
                
                .Text = "<<UN>>"
        
                .Replacement.Text = UN
                
                .Execute Replace:=wdReplaceAll
                
                .Text = "<<REASON>>"
        
                .Replacement.Text = Reason
                
                .Execute Replace:=wdReplaceAll
        
              End With
              
              'Get next linked story (if any)
        
              Set rngStory = rngStory.NextStoryRange
        
            Loop Until rngStory Is Nothing
            Next
    End If
    
    'MsgBox "Total time was: " & (Timer - startTime)
    
End Sub
