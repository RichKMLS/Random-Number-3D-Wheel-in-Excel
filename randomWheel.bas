Option Explicit

'---DECLARE PUBLIC VARIABLES---
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Public archiveDir, theSlice As String
Public archiveTotal As Integer
Public archiveValue As Boolean
Public WheelWB As Workbook
Public WheelWS As Worksheet
    
  
Sub archiveruns()

    Debug.Print "starting archiveruns()" & _
        vbNewLine & "-----------------------" & _
        vbNewLine

    
    '---DEFINE VARIABLES---
    
    'Total number of times to run in succession
    archiveTotal = 1
    
    'Directory where ArchiveResults.csv is stored.
        archiveDir = "C:\path\to\directory\"
    
    '----------------------
    
        If archiveDir <> "" and archiveDir <> "C:\path\to\directory\" Then
        
        archiveValue = True
        Call spinwheel
        
    Else
    
        MsgBox "Please set the directory where you wish to store ArchiveResults.csv"
            
    End If
    
    
End Sub
 
 

Sub spinwheel()

On Error Resume Next

With Application
    .Interactive = False
    .DisplayAlerts = False
End With

    '---DECLARE VARIABLES---
    
    Dim k, i, j, a, d, e As Integer
    Dim thenum, lastnum, thestartnum As Integer
    Dim suspenseRounds, suspenseSleep, suspenseOdds As Integer
        
    Dim wslice As Variant
    
    Dim thestart, initialspin, suspense As Boolean
    
    '---DEFINE VARIABLES---
    
    wslice = Array("c_1", "c_2", "c_3", "c_4", "c_5", _
    "c_6", "c_7", "c_8", "c_9", "c_10", "c_11", "c_12")
    
    k = 0
    thestartnum = -1
    
    e = 2000
    suspenseOdds = 100
    
    thestart = True
    initialspin = True

    '---DEFINE DEFAULT WB & WS---
    
    Set WheelWB = Workbooks("theAmazingSpinningWheel.xlsm")
    Set WheelWS = WheelWB.Worksheets("Sheet1")
    
    '----------------------
    
    Debug.Print "starting spinwheel()" & _
        vbNewLine & "-----------------------" & _
        vbNewLine
    
    With WheelWS
        
        .Range("$D$8").Select   ' hide selected cell
                                '   in case user clicked a cell
        DoEvents

        For a = 0 To 11 ' locate the current green slice
            If .Shapes.Range(Array(wslice(a))).Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent6 Then
                thestartnum = a
                DoEvents
                Exit For
            End If
        Next a
    
        DoEvents
    
        For j = 0 To 10000
    
            For i = 0 To 11
    
                Application.ScreenUpdating = False  ' turn screenupdating off
                                                    '   while new colors are assigned
                If thestart = True Then
                    If thestartnum <> -1 Then
                        i = thestartnum + 1         ' start the spin at the current
                        DoEvents                    '   green slice defined earlier
                        Debug.Print "start updated to slice #" & i
                    End If
                    thestart = False
                    DoEvents
                    Sleep 20
                End If
    
                If i = 12 Then i = 0
    
                thenum = i + 1
                lastnum = thenum - 1
                If thenum = 1 Or thenum = 13 Then
                    thenum = 1
                    lastnum = 12
                End If
    
                'define the font color for this slice and reset the previous slice
                .Shapes.Range(Array("num_" & thenum)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Shapes.Range(Array("num_" & lastnum)).TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
                DoEvents
    
                ' update the new slice to be selected (grey color)
                '   reset the previous slice to red or black
                With .Shapes.Range(Array(wslice(i))).Fill
                    If i Mod 2 <> False Then
    
                        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                        .ForeColor.Brightness = -0.2
    
                        WheelWS.Shapes.Range(Array(wslice(i - 1))).Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
                        WheelWS.Shapes.Range(Array(wslice(i - 1))).Fill.ForeColor.TintAndShade = 0
                        WheelWS.Shapes.Range(Array(wslice(i - 1))).Fill.ForeColor.Brightness = 0.150000006
    
                    Else
                        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                        .ForeColor.Brightness = -0.3
    
                        If i <> 0 Then
                            WheelWS.Shapes.Range(Array(wslice(i - 1))).Fill.ForeColor.RGB = RGB(208, 70, 56)
                        Else
                            WheelWS.Shapes.Range(Array("c_12")).Fill.ForeColor.RGB = RGB(208, 70, 56)
                        End If
    
                    End If
    
                End With
    
                ' turn screenupdating on to display the newly selected slice
                Application.ScreenUpdating = True
                
                If suspense = True Then
                    
                    DoEvents
                    
                    d = randomNum(suspenseOdds, True)
                    
                    DoEvents
                    
                    Debug.Print d
                    
                    If d <> 1 Then
    
                        suspenseRounds = suspenseRounds + 1
                        
                        suspenseOdds = d
                        
                        DoEvents
                        
                        suspenseSleep = Int((k + 20) * (suspenseRounds * 1.4))
                    
                        If suspenseSleep > 6000 Then suspenseSleep = 6000
    
                        DoEvents
                        Sleep suspenseSleep
                        DoEvents
                        
                        Debug.Print "MORE SUSPENSE  -  " & suspenseRounds & "  ->  " & Int((k + 20) * (suspenseRounds * 1.33))
                        
                        'skip to the end of the sub to loop back for another suspense round.
                        GoTo thesuspense
                    Else
                        
                        suspenseSleep = Int((k + 20) * (suspenseRounds * 1.33))
                        
                        If suspenseSleep > 6000 Then suspenseSleep = 6000
                
                        DoEvents
                        Sleep Int(suspenseSleep * 1.5)
                        DoEvents
                        
                        theSlice = wslice(i)
                        
                        Call thewinner
                        Exit Sub
                        
                    End If
    
                End If

                If j > 2 Then
    
                    d = randomNum(e, False)
                    Debug.Print "randomNum: " & d
    
                    DoEvents

                    If d = 1 Then
                        
                        DoEvents
                        Sleep k
                        DoEvents
                        
                        If e < 200 Then
                            d = randomNum(20, True)
                        Else
                            ' Quick Round probability increased (no suspense)
                            d = randomNum(10, True)
                        End If
                        
                        DoEvents
                        
                        If d = 1 Then
                        
                            suspense = False
                            
                            Debug.Print "Skip Suspense"
                        
                        Else
                        
                            Debug.Print "SUSPENSE"
                            
                            suspense = True
                            
                            suspenseRounds = 1
                            
                            e = 1
                            
                            DoEvents
                            Sleep k + 10
                            DoEvents
                            
                        End If
    
                        If suspense = False Then
    
                            suspenseSleep = (k + 10) * suspenseRounds
                            
                            If suspenseSleep > 6000 Then suspenseSleep = 6000
                            
                            DoEvents
                            Sleep Int(suspenseSleep * 1.5)
                            DoEvents
                            
                            theSlice = wslice(i)
                            
                            Call thewinner
                            Exit Sub
                            
                        End If
                        
                    Else
    
                        d = randomNum(100, True)
                        
                        DoEvents
                        
                        e = e - d

                        DoEvents
    
                        If e < 1 Then e = 5
                        
                        Debug.Print "Current Odds: 1/" & e
    
                    End If
    
                    k = k + j
    
                End If
    
thesuspense:
    
                DoEvents
                Sleep 10 + k
                DoEvents
    
            Next i
            DoEvents
        Next j
    
        'This shouldn't have happened! 
        With Application
            .ScreenUpdating = True
            .Interactive = True
            .DisplayAlerts = True
        End With

    End With

End Sub



Sub thewinner()
' Displays a simple animation where the selected slice will flash green and light green
'   indicating that it is the winning slice. This slice will be used as the starting
'   slice of the next spin round.

With Application
    .Interactive = False
    .ScreenUpdating = True
End With
    
    Dim i As Integer

    DoEvents
    
    Debug.Print "starting thewinner()" & _
        vbNewLine & "-----------------------" & _
        vbNewLine

    With WheelWS.Shapes.Range(theSlice).Fill
    
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
    
        For i = 1 To 21
    
            DoEvents
    
            If i Mod 2 <> False Then
                .ForeColor.Brightness = -0.2
                DoEvents
            Else
                .ForeColor.Brightness = 0.6
                DoEvents
            End If
    
            DoEvents
            Sleep 100
            DoEvents
    
        Next i
        
    End With
    
    If archiveValue = True Then
        Call archiveResult
    Else
        Call funFinished
    End If
    
End Sub


Sub archiveResult()

    With Application
        .Interactive = False
        .DisplayAlerts = False
    End With
    
    Debug.Print "Starting archiveResult()" & _
        vbNewLine & "-----------------------" & _
        vbNewLine

    Dim sliceStr As String
    Dim theResult As String

    sliceStr = theSlice
    sliceStr = Split(sliceStr, "_")(1)

    theResult = WheelWS.Shapes("num_" & sliceStr).TextFrame.Characters.Text
    
    Open archiveDir & "ArchivedResults.csv" For Append As #1

        Print #1, theResult

    Close #1
    
    archiveTotal = archiveTotal - 1
    
    DoEvents
    
    If archiveTotal >= 0 Then

        Debug.Print "Archive runs remaining: " & archiveTotal
 
        If archiveTotal = 0 Then
            Call funFinished
            Exit Sub
        Else
            DoEvents
            Call spinwheel
        End If
        
    Else
    
        Call funFinished

    End If
    
End Sub

Sub funFinished()

Debug.Print "The fun has now concluded." & _
    vbNewLine & "-----------------------" & _
    vbNewLine
    
With Application
    .Interactive = True
    .DisplayAlerts = True
End With
        
End Sub



Function randomNum(outOf As Integer, randoutof As Boolean)
'Random number generator for generating a number from 1 to 'outOf'
    Randomize
    DoEvents
    randomNum = Int(1 + Rnd() * outOf)
    DoEvents
    
    ' if randoutof = True then this will take the randomNum generated and then generate
    '    another number using randomNum as the new upper bounds.
    If randoutof = True Then
        Randomize
        DoEvents
        randomNum = Int(1 + Rnd() * randomNum)
        DoEvents
    End If
    
End Function
