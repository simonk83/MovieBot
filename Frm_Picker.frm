VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Picker 
   Caption         =   "Choose Movie"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   OleObjectBlob   =   "Frm_Picker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Lbl_Movie3_Click()

End Sub

Private Sub Opt_Movie0_Click()
    ChosenMovie = 0
    ChosenName = Left(Frm_Picker.Lbl_Movie0, Len(Frm_Picker.Lbl_Movie0) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie1_Click()
    ChosenMovie = 1
    ChosenName = Left(Frm_Picker.Lbl_Movie1, Len(Frm_Picker.Lbl_Movie1) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie2_Click()
    ChosenMovie = 2
    ChosenName = Left(Frm_Picker.Lbl_Movie2, Len(Frm_Picker.Lbl_Movie2) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie3_Click()
    ChosenMovie = 3
    ChosenName = Left(Frm_Picker.Lbl_Movie3, Len(Frm_Picker.Lbl_Movie3) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie4_Click()
    ChosenMovie = 4
    ChosenName = Left(Frm_Picker.Lbl_Movie4, Len(Frm_Picker.Lbl_Movie4) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie5_Click()
    ChosenMovie = 5
    ChosenName = Left(Frm_Picker.Lbl_Movie5, Len(Frm_Picker.Lbl_Movie5) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie6_Click()
    ChosenMovie = 6
    ChosenName = Left(Frm_Picker.Lbl_Movie6, Len(Frm_Picker.Lbl_Movie6) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie7_Click()
    ChosenMovie = 7
    ChosenName = Left(Frm_Picker.Lbl_Movie7, Len(Frm_Picker.Lbl_Movie7) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie8_Click()
    ChosenMovie = 8
    ChosenName = Left(Frm_Picker.Lbl_Movie8, Len(Frm_Picker.Lbl_Movie8) - 7)
    Unload Me
End Sub
Private Sub Opt_Movie9_Click()
    ChosenMovie = 9
    ChosenName = Left(Frm_Picker.Lbl_Movie9, Len(Frm_Picker.Lbl_Movie9) - 7)
    Unload Me
End Sub
