VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   ClientHeight    =   10514
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14520
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_PRun_Click()
    
    'Convert to proper case
    For x = 0 To Form.Controls.Count - 1
    
        If Trim(Left(Form.Controls(x).Name, Len("Txt_Director"))) = "Txt_Director" Then
        
            Form.Controls("Txt_Director" & Replace(Form.Controls(x).Name, "Txt_Director", "")).Value = StrConv(Form.Controls("Txt_Director" & Replace(Form.Controls(x).Name, "Txt_Director", "")).Value, vbProperCase)
        
        End If
        
        If Trim(Left(Form.Controls(x).Name, Len("Txt_Actor"))) = "Txt_Actor" Then
          
            Form.Controls("Txt_Actor" & Replace(Form.Controls(x).Name, "Txt_Actor", "")).Value = StrConv(Form.Controls("Txt_Actor" & Replace(Form.Controls(x).Name, "Txt_Actor", "")).Value, vbProperCase)
        
        End If
        
    Next x
    
    Call ThisWorkbook.PersonDeets

End Sub

Private Sub CommandButton1_Click()

'************************************
'This section composes the final post
'************************************

    Dim strBody As String
    Dim doClip As DataObject


    ' strBody = Form.Txt_Director & vbCr & vbCr & "[*Stars:*]  "


    For x = 0 To Form.Controls.Count - 1

        If Trim(Left(Form.Controls(x).Name, Len("Txt_Actor"))) = "Txt_Actor" Then

            If Form.Controls("Chk_Actor" & Replace(Form.Controls(x).Name, "Txt_Actor", "")).Value = True Then
                strBody = Form.Controls(x).Value & ", " & strBody
            End If

        End If

    Next x

    strBody = Form.Txt_Director & vbCr & vbCr & "[*Stars:*]  " & strBody


    strBody = Left(strBody, Len(strBody) - 2)

    strBody = strBody & vbCr & vbCr

    strBody = strBody & Txt_Synopsis.Value & vbCr & vbCr
    strBody = strBody & Txt_IMDB.Value & vbCr & vbCr

    If Len(Txt_Trailer.Value) > 0 Then
        strBody = strBody & Txt_Trailer.Value
    End If


    Set doClip = New DataObject
    'Put sText into the DataObject
    doClip.SetText strBody
    'Put the data in the DataObject into the Clipboard
    doClip.PutInClipboard

    Form.Opt_TMDB.Value = True

End Sub

Private Sub CommandButton2_Click()

    For x = 0 To Form.Controls.Count - 1

        If Controls(x).Name <> "Txt_Search" And Controls(x).Name <> "Txt_MYear" Then
            If Left(Controls(x).Name, 3) = "Txt" Then
                Controls(x).Value = ""
            End If

            If Left(Controls(x).Name, 3) Like "Chk" Then
                Controls(x).Value = False
            End If
        End If
    Next x

    Call ThisWorkbook.Search

End Sub



Private Sub Image1_Click()

'**********************************************************
'This section clears all fields (click on the header image)
'**********************************************************

    For x = 0 To Form.Controls.Count - 1

        If Left(Controls(x).Name, 3) = "Txt" Then
            Controls(x).Value = ""
        End If

        If Left(Controls(x).Name, 3) Like "Chk" Then
            Controls(x).Value = False
        End If

        Form.WebBrowser1.Visible = False

    Next x

End Sub


Private Sub Opt_Dir_Click()

    Form.Txt_Director.BackColor = &HFFFFFF
    Form.Txt_Director.Locked = False

End Sub

Private Sub Opt_IMDB_Click()

'Shows the Year text box (and image) when IMDB is chosen
    Form.Txt_MYear.Visible = True
    Form.Lbl_Year.Visible = True

End Sub

Private Sub Opt_Join_Click()

    Form.Txt_Director.BackColor = &HFFFFFF
    Form.Txt_Director.Locked = False

End Sub

Private Sub Opt_Star_Click()

'Grey and lock out the director text box when the Stars option is chosen.  It's just easier.
    Form.Txt_Director.BackColor = &HE0E0E0
    Form.Txt_Director.Locked = True

End Sub

Private Sub Opt_Talks_Click()

    Form.Txt_Director.BackColor = &HFFFFFF
    Form.Txt_Director.Locked = False

End Sub

Private Sub Opt_TMDB_Click()

    Form.Txt_MYear.Visible = False
    Form.Lbl_Year.Visible = False

End Sub


Private Sub OptionButton1_Click()

    Form.Txt_Director.BackColor = &HFFFFFF
    Form.Txt_Director.Locked = False

End Sub


Private Sub UserForm_Initialize()

    Form.Txt_MYear.Visible = False
    Form.Lbl_Year.Visible = False

End Sub

