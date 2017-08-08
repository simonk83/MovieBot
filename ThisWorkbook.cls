VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Search()

    Dim xDOC As DOMDocument
    Dim XMLHttpRequest As XMLHTTP
    Dim XMLElement As IXMLDOMElement
    Dim response As String
    Dim URL As String
    Dim x As String
    Dim Farge As Variant
    Dim YearNum As Variant
    Dim actorNum As Integer
    Dim MYear As String
    Dim js As Object
    Dim JData As String
    Dim pid As String
    Dim pname As String
    Dim pid2 As String
    Dim pname2 As String
    Dim OrigPoster As String
    

    On Error Resume Next
    
    Form.WebBrowser1.Visible = False

    'Checks for empty search term (main search box only)
    If Len(Form.Txt_Search.Value) < 1 Then
        MsgBox "Ummm enter something, dingus"
        End
    End If

    'Uses IMDB data from http://www.omdbapi.com
        
        If Form.Opt_IMDB.Value = True Then
        
        '***************************
        'OMDB
        '***************************
        
        x = get_Movie_ID_OMDB(Form.Txt_Search.Value)

        URL = "http://www.omdbapi.com/?i=" & x & "&plot=full&r=json"
        
        'Opens link
        JData = GetWebSource(URL)

        'Parses JSON data
        Set js = JSON.parse(JData)

        'Sets person ID and name
        If nocrew = False Then

            pname = Trim(js.Item("Director"))
            
            If pname = "N/A" Then
                pname = ""
            End If

            'Gets the second director, if present
'            On Error Resume Next
'            If Trim(js.Item("casts").Item("crew")(2).Item("job")) = "Director" Then
'
'                pid2 = Trim(js.Item("casts").Item("crew")(2).Item("id"))
'                pname2 = Trim(js.Item("casts").Item("crew")(2).Item("name"))
'
'            End If

            'Director
            Form!Txt_Director.Value = pname

            'Director 2
'            If pid2 <> "" Then
'                get_Person_Info (pid2)
'                Form!Txt_Director.Value = Form!Txt_Director.Value & ", <a href=""" & nameurl & get_Person_Info(pid2) & """" & ">" & pname2 & "</a>"
'            End If

        End If

        'IMDB ID
        If Not IsNull(js.Item("imdbID")) Then
            Form!Txt_IMDB.Value = "http://www.imdb.com/title/" & Trim(js.Item("imdbID"))
        End If

        'Trailer
        'notrailer is set during the first step of the JSON parse
'        If notrailer = False Then
'            Form!Txt_Trailer.Value = "[*<a href=" & """" & "http://www.youtube.com/watch?v=" & Trim(js.Item("trailers").Item("youtube")(1).Item("source")) & """" & ">Trailer</a>*]"
'        Else
'        End If

        'Synopsis
        Form!Txt_Synopsis.Value = "['" & Trim(js.Item("Plot")) & "']"

        'cast
        actorNum = 1

            pname = Trim(js.Item("Actors"))
            
            ActorList = Split(pname, ",")
            
            For i = 0 To UBound(ActorList)
                Form.Controls("txt_Actor" & actorNum) = Trim(ActorList(i))
                actorNum = actorNum + 1
                
                If actorNum > 10 Then
                 GoTo fiush
                End If
            Next i

fiush:

    'Poster
    OrigPoster = Trim(js.Item("Poster"))
    PosterURL = Replace(OrigPoster, "SX300", "SX154")


        'Can't make the webbrowser control invisible or the page won't load.  This does the trick
        Form.WebBrowser1.Left = -Form.WebBrowser1.Width

        Form.WebBrowser1.Visible = True

        Form.WebBrowser1.Navigate PosterURL

        PauseTime = 30    ' Set duration to wait in seconds.
        Start = Timer    ' Set start time.

        Do    ' While Timer < Start + PauseTime
            DoEvents    ' allow other processes to work (the browser to have time to load the webpage)
        Loop Until Form.WebBrowser1.readyState = READYSTATE_COMPLETE Or Timer > Start + PauseTime

        Form.WebBrowser1.Document.body.Scroll = "no"
        Form.WebBrowser1.Document.body.Style.Border = "none"

        Form.WebBrowser1.Left = 204


    'Uses IMDB data from http://mymovieapi.com
'    If Form.Opt_IMDB.Value = True Then
'
'        MYear = Form.Txt_MYear.Value
'
'        URL = "http://mymovieapi.com/?title=" & Form.Txt_Search.Value & "&type=xml&plot=full&episode=0&limit=10"
'
'        If Len(MYear) > 0 Then
'            URL = URL & "&year=" & MYear & "&yg=1"
'        Else
'            URL = URL & "&yg=0"
'        End If
'
'
'TryAgain:
'        'Opens URL
'        Set XMLHttpRequest = New MSXML2.XMLHTTP
'        With XMLHttpRequest
'            .Open "GET", URL, False
'            .send
'        End With
'
'
'        'Probably only useful for works crappy network
'        If XMLHttpRequest.Status = 500 Then    'goto tryagain
'            MsgBox ("Network issue")
'            End
'        End If
'
'
'        Set xDOC = New DOMDocument
'        Do Until xDOC.readyState = 4
'        Loop
'        xDOC.loadXML (XMLHttpRequest.responseText)
'
'
'        'Checks for multiple title matches
'        Set Farge = xDOC.getElementsByTagName("Title")
'
'        movienum = 0
'
'        For Each XMLElement In Farge
'
'            Frm_Picker.Controls("Lbl_Movie" & movienum) = XMLElement.Text
'            movienum = movienum + 1
'
'        Next XMLElement
'
'
'        'Grabs year of release for each title
'        Set Farge = xDOC.getElementsByTagName("year")
'
'        'Sets the width of the labels according to the text length.  Dirty way of doing it, but why break tradition
'        YearNum = 0
'        maxwidth = Frm_Picker.Controls("Lbl_Movie0").Width
'
'        For Each XMLElement In Farge
'
'            Frm_Picker.Controls("Lbl_Movie" & YearNum) = Frm_Picker.Controls("Lbl_Movie" & YearNum) & " (" & XMLElement.Text & ")"
'
'            If Len(Frm_Picker.Controls("Lbl_Movie" & YearNum)) > 10 Then Frm_Picker.Controls("Lbl_Movie" & YearNum).Width = Len(Frm_Picker.Controls("Lbl_Movie" & YearNum)) * 7
'
'            lblwidth = Frm_Picker.Controls("Lbl_Movie" & YearNum).Width
'
'            If lblwidth > maxwidth Then
'                maxwidth = lblwidth
'            End If
'
'            YearNum = YearNum + 1
'
'        Next XMLElement
'
'
'        'Sets all label widths to the largest one
'        For y = 0 To Frm_Picker.Controls.Count - 1
'
'            If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then
'
'                Frm_Picker.Controls(y).Width = maxwidth
'
'            End If
'
'        Next y
'
'
'        'If there are multiple options, show the form, otherwise...don't
'        If movienum > 1 Then
'
'            Load Frm_Picker
'
'            'Sets the form width based on the largest label width, plus a bit
'            Frm_Picker.Width = maxwidth * 1.2
'
'            For y = 0 To Frm_Picker.Controls.Count - 1
'
'                If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then
'
'                    'Hides any unused controls and resizes form height, approximately
'                    If Frm_Picker.Controls(y) = "" Then
'                        Frm_Picker.Controls("Opt_Movie" & Right(Frm_Picker.Controls(y).Name, 1)).Visible = False
'                        Frm_Picker.Height = Frm_Picker.Height - 36
'                        Frm_Picker.Image2.Top = Frm_Picker.Image2.Top - 36
'                    End If
'
'                End If
'
'            Next y
'
'            'Show the form
'            Frm_Picker.Show
'
'        Else
'
'            'If there's only 1 option, choose that!
'            ChosenMovie = 0
'            ChosenName = Form.Txt_Search.Value
'
'        End If
'
'        'Collects IMDB ID
'        Set Farge = xDOC.getElementsByTagName("imdb_id")
'
'        If xDOC.Text = "404 Film not found" Then
'            MsgBox "No IMDB ID for " & ChosenName
'            End
'        End If
'
'        'Adds ID to form
'        Form!Txt_IMDB.Value = "http://www.imdb.com/title/" & Farge.Item(ChosenMovie).Text
'
'        'Switches to use the exact movie we chose earlier (using IMDB ID).  Avoids out of order items in the XML
'        URL = "http://mymovieapi.com/?id=" & Farge.Item(ChosenMovie).Text & "&type=xml&plot=full&episode=0"
'
'        'Opens URL
'        Set XMLHttpRequest = New MSXML2.XMLHTTP
'        With XMLHttpRequest
'            .Open "GET", URL, False
'            .send
'        End With
'
'        'Probably only useful for works crappy network
'        If XMLHttpRequest.Status = 500 Then    'goto tryagain
'            MsgBox ("Network issue")
'            End
'        End If
'
'        Set xDOC = New DOMDocument
'        Do Until xDOC.readyState = 4
'        Loop
'        xDOC.loadXML (XMLHttpRequest.responseText)
'
'
'        'Collects synopsis
'        Set Farge = xDOC.getElementsByTagName("plot")
'
'        If Not Farge.Length = "0" Then
'            Form!Txt_Synopsis.Value = "['" & Farge.Item(0).Text & "']"
'        End If
'
'        'Collects director name
'        Set Farge = xDOC.getElementsByTagName("directors")
'
'        'Checks for no director and skips
'        If Farge.Length = 0 Then
'            GoTo skipdir
'        End If
'
'        'Checks for multiple directors.  Will just add both with a comma separating.  No hyperlink gathered (yet)
'        If Farge.Item(0).childNodes.Length > 1 Then
'
'            For Each XMLElement In Farge.Item(0).childNodes
'
'                Form!Txt_Director.Value = Form!Txt_Director.Value & XMLElement.Text & ", "
'
'            Next XMLElement
'
'            Form!Txt_Director.Value = Left(Form!Txt_Director.Value, Len(Form!Txt_Director.Value) - 2)
'
'        Else
'
'            Form!Txt_Director.Value = Farge.Item(0).Text
'
'        End If

'skipdir:
'
'        'Collects actor names
'        Set Farge = xDOC.getElementsByTagName("actors")
'
'        If Farge.Length = 0 Then
'            Call PersonDeets
'            Exit Sub
'        End If
'
'        actorNum = 1
'
'        For Each XMLElement In Farge.Item(0).childNodes
'
'            'Only gather 10 actors max
'            If actorNum > 10 Then
'                GoTo Flumps
'            End If
'
'            Form.Controls("txt_Actor" & actorNum).Value = XMLElement.Text
'            actorNum = actorNum + 1
'
'        Next XMLElement
'
'Flumps:
'
        Call PersonDeets


    Else

        '*****************************
        'TMDB
        '*****************************

        'Uses TMDB data - API 3 - JSON
        x = get_Movie_ID(Form.Txt_Search.Value)

        URL = "http://api.themoviedb.org/3/movie/" & x & "?api_key=d7510a5d85a6e71f4df80f8e49aba9e7&append_to_response=casts,trailers"

        'Opens link
        JData = GetWebSource(URL)

        'Parses JSON data
        Set js = JSON.parse(JData)

        'Sets person ID and name
        If nocrew = False Then
            
            j = 1
            
            Do Until Trim(js.Item("casts").Item("crew")(j).Item("job")) = "Director"

                j = j + 1
                
            Loop
            
            pid = Trim(js.Item("casts").Item("crew")(j).Item("id"))
            pname = Trim(js.Item("casts").Item("crew")(j).Item("name"))
            
            j = j + 1
            
            'Gets the second director, if present
            On Error GoTo nope
            If Trim(js.Item("casts").Item("crew")(j).Item("job")) = "Director" Then

                pid2 = Trim(js.Item("casts").Item("crew")(j).Item("id"))
                pname2 = Trim(js.Item("casts").Item("crew")(j).Item("name"))

            End If

nope:
            'Director
            Form!Txt_Director.Value = "[*Director:*]  <a href=" & """"
            get_Person_Info (pid)
            Form!Txt_Director.Value = Form!Txt_Director.Value & nameurl & get_Person_Info(pid) & """" & ">" & pname & "</a>"

            'Director 2
            If pid2 <> "" Then
                get_Person_Info (pid2)
                Form!Txt_Director.Value = Form!Txt_Director.Value & ", <a href=""" & nameurl & get_Person_Info(pid2) & """" & ">" & pname2 & "</a>"
            End If

        End If

        'IMDB ID
        If Not IsNull(js.Item("imdb_id")) Then
            Form!Txt_IMDB.Value = "http://www.imdb.com/title/" & Trim(js.Item("imdb_id"))
        End If

        'Trailer
        'notrailer is set during the first step of the JSON parse
        If notrailer = False Then
            Form!Txt_Trailer.Value = "[*<a href=" & """" & "http://www.youtube.com/watch?v=" & Trim(js.Item("trailers").Item("youtube")(1).Item("source")) & """" & ">Trailer</a>*]"
        Else
        End If

        'Synopsis
        Form!Txt_Synopsis.Value = "['" & Trim(js.Item("overview")) & "']"

        'cast
        actorNum = 1

        For Each Item In js.Item("casts").Item("cast")

            If actorNum > 10 Then
                GoTo ActorMax
            End If

            pid = Trim(js.Item("casts").Item("cast")(actorNum).Item("id"))
            pname = Trim(js.Item("casts").Item("cast")(actorNum).Item("name"))

            Form.Controls("txt_Actor" & actorNum).Value = "<a href=" & """"
            get_Person_Info (pid)
            Form.Controls("txt_Actor" & actorNum).Value = Form.Controls("txt_Actor" & actorNum).Value & nameurl & get_Person_Info(pid) & """" & ">" & pname & "</a>"

            actorNum = actorNum + 1
            nameurl = ""

        Next Item

    End If

ActorMax:

    'Poster
    If PosterURL <> "" Then

        'Can't make the webbrowser control invisible or the page won't load.  This does the trick
        Form.WebBrowser1.Left = -Form.WebBrowser1.Width

        Form.WebBrowser1.Visible = True

        Form.WebBrowser1.Navigate PosterURL

        PauseTime = 30    ' Set duration to wait in seconds.
        Start = Timer    ' Set start time.

        Do    ' While Timer < Start + PauseTime
            DoEvents    ' allow other processes to work (the browser to have time to load the webpage)
        Loop Until Form.WebBrowser1.readyState = READYSTATE_COMPLETE Or Timer > Start + PauseTime

        Form.WebBrowser1.Document.body.Scroll = "no"
        Form.WebBrowser1.Document.body.Style.Border = "none"

        Form.WebBrowser1.Left = 204

    End If

End Sub


Function get_Movie_ID(SearchTerm As String) As String

    Dim URL As String
    Dim js As Object
    Dim JData As String

    URL = "http://api.themoviedb.org/3/search/movie?api_key=d7510a5d85a6e71f4df80f8e49aba9e7&query=" & Form.Txt_Search.Value

    JData = GetWebSource(URL)

    Set js = JSON.parse(JData)

    'Checks the Total Results parameter
    If Right(JData, 3) = ":0}" Then
        MsgBox ("Nothing Found for ") & SearchTerm
        End
    End If

    movienum = 1

    'Get the first 5 movie titles/years
    For Each Item In js.Item("results")

        If movienum > 10 Then
            GoTo done
        End If

        'Assigns each movie to a spot on the form
        Frm_Picker.Controls("Lbl_Movie" & movienum - 1) = js.Item("results")(movienum).Item("title") & " (" & Left(js.Item("results")(movienum).Item("release_date"), 4) & ")"

        movienum = movienum + 1

    Next Item

done:

    'Sets the width of the labels according to the text size.  Dirty way of doing it, but why break tradition
    maxwidth = Frm_Picker.Controls("Lbl_Movie0").Width
    LblNum = 0

    For y = 0 To Frm_Picker.Controls.Count - 1

        If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then

            If Len(Frm_Picker.Controls("Lbl_Movie" & LblNum)) > 5 Then Frm_Picker.Controls("Lbl_Movie" & LblNum).Width = Len(Frm_Picker.Controls("Lbl_Movie" & LblNum)) * 7

            lblwidth = Frm_Picker.Controls("Lbl_Movie" & LblNum).Width

            If lblwidth > maxwidth Then
                maxwidth = lblwidth
            End If

            LblNum = LblNum + 1

        End If
    Next y


    'Sets all label widths to the largest one
    For y = 0 To Frm_Picker.Controls.Count - 1

        If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then

            Frm_Picker.Controls(y).Width = maxwidth

        End If

    Next y


    'If there are multiple options, show the form, otherwise...don't
    If movienum > 2 Then

        Load Frm_Picker

        'Sets the form width based on the largest label width, plus a bit
        If Frm_Picker.Width < maxwidth * 1.18 Then

            Frm_Picker.Width = maxwidth * 1.18

        Else

            Frm_Picker.Width = Frm_Picker.Width

        End If

        For y = 0 To Frm_Picker.Controls.Count - 1

            If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then

                'Hides any unused controls and resizes form height, approximately
                If Frm_Picker.Controls(y) = "" Then
                    Frm_Picker.Controls("Opt_Movie" & Right(Frm_Picker.Controls(y).Name, 1)).Visible = False
                    Frm_Picker.Height = Frm_Picker.Height - 36
                    Frm_Picker.Image2.Top = Frm_Picker.Image2.Top - 36
                End If

            End If

        Next y

        Frm_Picker.Show

    Else

        ChosenMovie = 0

    End If


    'This is just because the form labels start at 0 (helps with the IMDB scrape) and the JSON starts at 1.  Lovely and neat!
    ChosenMovie = ChosenMovie + 1


    'Grab the poster
    If Not IsNull(js.Item("results")(ChosenMovie).Item("poster_path")) Then

        'PosterURL = "https://d3gtl9l2a4fn1j.cloudfront.net/t/p/w154" & Trim(js.Item("results")(ChosenMovie).Item("poster_path"))
        PosterURL = "https://image.tmdb.org/t/p/w154" & Trim(js.Item("results")(ChosenMovie).Item("poster_path"))
        
    Else

        PosterURL = ""

    End If


    get_Movie_ID = Trim(js.Item("results")(ChosenMovie).Item("id"))


End Function
Function get_Movie_ID_OMDB(SearchTerm As String) As String

    Dim URL As String
    Dim js As Object
    Dim JData As String

    URL = "http://www.omdbapi.com/?s=" & Form.Txt_Search.Value

    JData = GetWebSource(URL)

    Set js = JSON.parse(JData)

    'Checks the Total Results parameter
'    If Right(JData, 3) = ":0}" Then
'        MsgBox ("Nothing Found for ") & SearchTerm
'        End
'    End If

    movienum = 1

    'Get the first 10 movie titles/years
    For Each Item In js.Item("Search")

        If movienum > 10 Then
            GoTo done
        End If

        'Assigns each movie to a spot on the form
        Frm_Picker.Controls("Lbl_Movie" & movienum - 1) = js.Item("Search")(movienum).Item("Title") & " (" & js.Item("Search")(movienum).Item("Year") & ")"

        movienum = movienum + 1

    Next Item

done:

    'Sets the width of the labels according to the text size.  Dirty way of doing it, but why break tradition
    maxwidth = Frm_Picker.Controls("Lbl_Movie0").Width
    LblNum = 0

    For y = 0 To Frm_Picker.Controls.Count - 1

        If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then

            If Len(Frm_Picker.Controls("Lbl_Movie" & LblNum)) > 5 Then Frm_Picker.Controls("Lbl_Movie" & LblNum).Width = Len(Frm_Picker.Controls("Lbl_Movie" & LblNum)) * 7

            lblwidth = Frm_Picker.Controls("Lbl_Movie" & LblNum).Width

            If lblwidth > maxwidth Then
                maxwidth = lblwidth
            End If

            LblNum = LblNum + 1

        End If
    Next y


    'Sets all label widths to the largest one
    For y = 0 To Frm_Picker.Controls.Count - 1

        If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then

            Frm_Picker.Controls(y).Width = maxwidth

        End If

    Next y


    'If there are multiple options, show the form, otherwise...don't
    If movienum > 2 Then

        Load Frm_Picker

        'Sets the form width based on the largest label width, plus a bit
        If Frm_Picker.Width < maxwidth * 1.18 Then

            Frm_Picker.Width = maxwidth * 1.18

        Else

            Frm_Picker.Width = Frm_Picker.Width

        End If

        For y = 0 To Frm_Picker.Controls.Count - 1

            If Left(Frm_Picker.Controls(y).Name, 3) = "Lbl" Then

                'Hides any unused controls and resizes form height, approximately
                If Frm_Picker.Controls(y) = "" Then
                    Frm_Picker.Controls("Opt_Movie" & Right(Frm_Picker.Controls(y).Name, 1)).Visible = False
                    Frm_Picker.Height = Frm_Picker.Height - 36
                    Frm_Picker.Image2.Top = Frm_Picker.Image2.Top - 36
                End If

            End If

        Next y

        Frm_Picker.Show

    Else

        ChosenMovie = 0

    End If


    'This is just because the form labels start at 0 (helps with the IMDB scrape) and the JSON starts at 1.  Lovely and neat!
    ChosenMovie = ChosenMovie + 1



    get_Movie_ID_OMDB = Trim(js.Item("Search")(ChosenMovie).Item("imdbID"))


End Function


Public Function DisplayNode(ByRef Nodes As IXMLDOMNodeList)

'Currently not used ***************************************

    Dim xNode As IXMLDOMNode
    Dim Start As String
    Dim Finish As String
    Dim Output As String


    For Each xNode In Nodes
        If xNode.NodeType = NODE_CDATA_SECTION And _
           InStr(xNode.NodeValue, "id:") <> 0 Then
            Start = InStr(xNode.NodeValue, "Temperature:") + 17
            Finish = InStr(xNode.NodeValue, Chr(35)) - 1
            DisplayNode = Mid(xNode.NodeValue, Start, Finish - Start)
        End If
        If DisplayNode <> "" Then Exit Function
        If xNode.HasChildNodes Then
            Output = DisplayNode(xNode.childNodes)
            If Output <> "" Then DisplayNode = Output
        End If
    Next xNode

End Function

Private Sub Workbook_Open()

'Hides the Year search box and label on open
    Form.Txt_MYear.Visible = False
    Form.Lbl_Year.Visible = False
    Form.Show

End Sub

Sub PersonDeets()

'*****************************************************************************************
'This section parses out individual people.  Used both for getting individual [*Stars:*],
'[*Director:*] etc plus also to resolve names when using the IMDB API
'*****************************************************************************************

    Dim URL As String
    Dim URL2 As String
    Dim x As String
    Dim w As Integer
    Dim Farge As Variant
    Dim NameHold As String
    Dim strBody As String
    Dim doClip As DataObject
    Dim addit As String
    Dim js As Object
    Dim JData As String
    Dim js2 As Object
    Dim JData2 As String
    Dim pid As String
    Dim pid2 As String
    Dim pname As String
    Dim personnum As Integer
    Dim biog As String

    For y = 0 To Form.Controls.Count - 1
        Debug.Print Form.Controls(y).Name
        If Left(Form.Controls(y).Name, 9) = "Txt_Actor" Or Left(Form.Controls(y).Name, 12) = "Txt_Director" Then

            If Len(Form.Controls(y).Value) > 0 Then

                'Checks for multiple directors and doesn't attempt to find their links
                If InStr(1, Form.Controls(y).Value, ",") Then
                    GoTo skip
                Else
                End If

                'Sets name to x
                x = ConvertAccent(Form.Controls(y).Value)
                NameHold = Form.Controls(y).Value

                URL = "http://api.themoviedb.org/3/search/person?api_key=d7510a5d85a6e71f4df80f8e49aba9e7&query=" & x

                JData = GetWebSource(URL)

                If InStr(1, JData, "total_results"":0") > 0 Then '   Right(JData, 17) = "total_results"":0}" Then
                    MsgBox ("Nothing Found for ") & NameHold
                    GoTo skip
                End If

                'Removes the } from the string to be able to count the result number
                'JData = Left(JData, Len(JData) - 17)

                'If Right(JData, 1) > 1 Then
                If Mid(JData, 27, 1) > 1 Then
                
                    Set js = JSON.parse(JData)

                    'Resultnum = Right(JData, 1)
                    
                    Resultnum = Mid(JData, 27, 1)
                    
                    personnum = 1


                    For w = 1 To Resultnum

                        If w = 5 Then
                            GoTo skipit
                        End If

                        If Not IsNull(Trim(js.Item("results")(w).Item("profile_path"))) Then
                            
                            'Frm_PersonPicker.Controls("WebBrowser" & personnum).Navigate "http://d3gtl9l2a4fn1j.cloudfront.net/t/p/w130" & Trim(js.Item("results")(w).Item("profile_path"))
                            Frm_PersonPicker.Controls("WebBrowser" & personnum).Navigate "https://image.tmdb.org/t/p/w130" & Trim(js.Item("results")(w).Item("profile_path"))

                            PauseTime = 10        ' Set duration to wait in seconds.
                            Start = Timer        ' Set start time.

                            Do        ' While Timer < Start + PauseTime
                                DoEvents        ' allow other processes to work (the browser to have time to load the webpage)
                            Loop Until Frm_PersonPicker.Controls("WebBrowser" & personnum).readyState = READYSTATE_COMPLETE Or Timer > Start + PauseTime
                            
                            Frm_PersonPicker.Controls("WebBrowser" & personnum).Document.body.Scroll = "no"
                            Frm_PersonPicker.Controls("WebBrowser" & personnum).Document.body.Style.Border = "none"

                        Else

                            Frm_PersonPicker.Controls("WebBrowser" & personnum).Navigate "http://1.bp.blogspot.com/_ky1bf81QrMw/TUlSgZKc0vI/AAAAAAAABA0/K4ClLDL5opM/s150/no_photo_male.jpg"
                            
                            PauseTime = 10        ' Set duration to wait in seconds.
                            Start = Timer        ' Set start time.

                            Do        ' While Timer < Start + PauseTime
                                DoEvents        ' allow other processes to work (the browser to have time to load the webpage)
                            Loop Until Frm_PersonPicker.Controls("WebBrowser" & personnum).readyState = READYSTATE_COMPLETE Or Timer > Start + PauseTime
                            
                            Frm_PersonPicker.Controls("WebBrowser" & personnum).Document.body.Scroll = "no"
                            Frm_PersonPicker.Controls("WebBrowser" & personnum).Document.body.Style.Border = "none"
                            
                        End If


                        pid2 = Trim(js.Item("results")(personnum).Item("id"))
                        URL2 = "http://api.themoviedb.org/3/person/" & pid2 & "?api_key=d7510a5d85a6e71f4df80f8e49aba9e7"

                        JData2 = GetWebSource(URL2)

                        Set js2 = JSON.parse(JData2)

                        If Not IsNull(Trim(js2.Item("biography"))) Then

                            Frm_PersonPicker.Controls("Textbox" & personnum) = Trim(js2.Item("biography"))
                            Frm_PersonPicker.Controls("Textbox" & personnum).Visible = True

                        End If


                        personnum = personnum + 1


                    Next w

skipit:

                    Frm_PersonPicker.Show

                Else

                    ChosenPerson = 1

                End If
                
                

                Set js = JSON.parse(JData)

                pid = Trim(js.Item("results")(ChosenPerson).Item("id"))


                'If the Stars option is checked, or the IMDB option, don't bold the name
                If Form.Opt_Star.Value = True Or Form.Opt_IMDB.Value = True Then

                    If Form.Controls(y).Name = "Txt_Director" Then

                        Form.Controls(y).Value = "[*Director:*] <a href=" & """"
                        get_Person_Info (pid)
                        Form.Controls(y).Value = Form.Controls(y).Value & nameurl & get_Person_Info(pid) & """" & ">" & NameHold & "</a>"

                    Else

                        Form.Controls(y).Value = "<a href=" & """"
                        get_Person_Info (pid)
                        Form.Controls(y).Value = Form.Controls(y).Value & nameurl & get_Person_Info(pid) & """" & ">" & NameHold & "</a>"

                    End If


                Else

                    Form.Controls(y).Value = "[*<a href=" & """"
                    get_Person_Info (pid)
                    Form.Controls(y).Value = Form.Controls(y).Value & nameurl & get_Person_Info(pid) & """" & ">" & NameHold & "</a>*]"

                End If

                'strBody = strBody & Form.Controls(y).Value & ", "
                strBody = Form.Controls(y).Value & ", " & strBody

            End If

        End If

        NameHold = ""

skip:

        notfound = False

    Next y

    If Len(strBody) > 0 Then

        If Form.Opt_IMDB.Value = False Then

            'Removes the final comma and space

            strBody = Left(strBody, Len(strBody) - 2)


            If Form.Opt_Join.Value = True Then

                'Checks for multiple people and gets the English all proper like
                If InStr(1, strBody, ",") Then
                    addit = " join the cast."
                Else
                    addit = " joins the cast."
                End If

                strBody = strBody & addit

                'Finds the last comma, and replaces it with "and"
                If InStr(1, strBody, ",") Then
                    strBody = Left(strBody, InStrRev(strBody, ",") - 1) & " and" & Right(strBody, Len(strBody) - InStrRev(strBody, ","))
                Else
                End If

            End If


            If Form.Opt_Dir.Value = True Then

                'If Director option chosen, remove prior formatting and add it back properly here
                strBody = Left(strBody, Len(strBody) - 2)
                strBody = Right(strBody, Len(strBody) - 2)
                strBody = "[*Director:*] " & strBody

            End If


            If Form.Opt_Star.Value = True Then

                strBody = "[*Stars:*] " & strBody

            End If


            If Form.Opt_Talks.Value = True Then

                strBody = strBody & " in talks."

                If InStr(1, strBody, ",") Then
                    strBody = Left(strBody, InStrRev(strBody, ",") - 1) & " and" & Right(strBody, Len(strBody) - InStrRev(strBody, ","))
                Else
                End If

            End If

        End If

        Set doClip = New DataObject
        'Put sText into the DataObject
        doClip.SetText strBody
        'Put the data in the DataObject into the Clipboard
        doClip.PutInClipboard

        GoTo Fin

    Else
    End If

    If Form.Txt_Search.Value <> "" Then
        'Poster
        PosterURL = get_TMDB_Poster(Form.Txt_Search.Value)
    
        If PosterURL <> "" Then
    
            Form.WebBrowser1.Left = -Form.WebBrowser1.Width
    
            Form.WebBrowser1.Visible = True
    
            Form.WebBrowser1.Navigate PosterURL
    
            PauseTime = 30        ' Set duration to wait in seconds.
            Start = Timer        ' Set start time.
    
            Do        ' While Timer < Start + PauseTime
                DoEvents        ' allow other processes to work (the browser to have time to load the webpage)
            Loop Until Form.WebBrowser1.readyState = READYSTATE_COMPLETE Or Timer > Start + PauseTime
    
            Form.WebBrowser1.Document.body.Scroll = "no"
            Form.WebBrowser1.Document.body.Style.Border = "none"
    
            Form.WebBrowser1.Left = 204
    
        End If
    End If
Fin:

    'Resets back to use TMDB
    Form.Opt_TMDB.Value = True

End Sub

Function get_Person_ID(SearchTerm As String) As String

    Dim xDOC As DOMDocument
    Dim XMLHttpRequest As XMLHTTP
    Dim response As String
    Dim URL As String

    'Calls the special character replace function to remove accents etc from names (XML isn't a fan)
    SearchTerm = ConvertAccent(SearchTerm)

    URL = "http://api.themoviedb.org/2.1/Person.search/en/xml/d7510a5d85a6e71f4df80f8e49aba9e7/" & SearchTerm


TryAgain:
    Set XMLHttpRequest = New MSXML2.XMLHTTP
    With XMLHttpRequest
        .Open "GET", URL, False
        .send
    End With
    

    If XMLHttpRequest.Status = 500 Then 'goto tryagain
        MsgBox ("Network issue")
        End
    End If
    
 
    Set xDOC = New DOMDocument
    Do Until xDOC.readyState = 4
    Loop
    xDOC.loadXML (XMLHttpRequest.responseText)
    

    'Checks for negative result
    If XMLHttpRequest.responseText = "<h1>Not Found</h1>" Then
        MsgBox "Nothing Found for " & SearchTerm
        End
    End If

    'Checks for negative result
    If xDOC.childNodes(1).Text = "0 Nothing found." Then
        MsgBox "Nothing Found for " & SearchTerm
        notfound = True
        GoTo skip
    End If


    get_Person_ID = xDOC.getElementsByTagName("id").Item(0).Text

skip:

End Function


Function ConvertAccent(ByVal inputString As String) As String
' http://www.vbforums.com/archive/index.php/t-483965.html

'**********************************************************************************
'Replaces any accented characters with their equivalents, otherwise MSXML has a cry
'**********************************************************************************

    Dim x As Long, Position As Long

    Const AccChars As String = _
          "¿¡¬√ƒ≈«»… ÀÃÕŒœ–—“”‘’÷Ÿ⁄€‹›‡·‚„‰ÂÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˘˙˚¸˝ˇ"
    Const RegChars As String = _
          "AAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"
    For x = 1 To Len(inputString)
        Position = InStr(AccChars, Mid(inputString, x, 1))
        If Position Then Mid(inputString, x) = Mid(RegChars, Position, 1)
    Next
    ConvertAccent = inputString
    
End Function


Public Function GetWebSource(ByRef URL As String) As String
'Opens URL's

    Dim xml As IXMLHTTPRequest

    On Error Resume Next
    Set xml = CreateObject("Microsoft.XMLHTTP")
    With xml
        .Open "GET", URL, False
        .send
        GetWebSource = .responseText
    End With
    Set xml = Nothing

End Function


Function get_Person_Info(SearchTerm As String) As String

    Dim URL As String
    Dim js As Object
    Dim JData As String

    URL = "http://api.themoviedb.org/3/person/" & SearchTerm & "?api_key=d7510a5d85a6e71f4df80f8e49aba9e7"

    JData = GetWebSource(URL)

    Set js = JSON.parse(JData)

    If IsNull(Trim(js.Item("imdb_id"))) Or Trim(js.Item("imdb_id")) = "" Then
    
        nameurl = "http://www.themoviedb.org/person/"
        get_Person_Info = SearchTerm
        
    Else
    
        nameurl = "http://www.imdb.com/name/"
        get_Person_Info = Trim(js.Item("imdb_id"))
        
    End If
    
    
End Function


Function get_TMDB_Poster(SearchTerm As String) As String

'***********************************************************************************************
'This is used purely because the posters from IMDB are too big, and it's not easy to resize them
'without saving them somewhere first.  This function basically grabs the name from the option
'you've chosen and tries to match it to a TMDB poster.  Not 100% accurate, but it'll do for now.
'***********************************************************************************************

    Dim URL As String
    Dim js As Object
    Dim JData As String
    Dim movienum As Integer
    Dim gotposter As Boolean
    
    URL = "http://api.themoviedb.org/3/search/movie?api_key=d7510a5d85a6e71f4df80f8e49aba9e7&query=" & Form.Txt_Search.Value

    JData = GetWebSource(URL)

    Set js = JSON.parse(JData)

    gotposter = False
    movienum = 1

    For Each Item In js.Item("results")

        Debug.Print js.Item("results")(movienum).Item("title")

        'ChosenName is set during the frm_picker process
        'Converts both to uppercase to avoid case issues
        If Replace(UCase(js.Item("results")(movienum).Item("title")), ".", "") = UCase(ChosenName) Then
            gotposter = True
            GoTo gotthesucker
        End If

        movienum = movienum + 1
    Next Item

    If gotposter = False Then
        get_TMDB_Poster = ""
        GoTo skip
    End If


gotthesucker:

    If Right(JData, 3) <> ":0}" Then

        If Not IsNull(js.Item("results")(movienum).Item("poster_path")) Then

            'get_TMDB_Poster = "https://d3gtl9l2a4fn1j.cloudfront.net/t/p/w154" & Trim(js.Item("results")(movienum).Item("poster_path"))
            get_TMDB_Poster = "https://image.tmdb.org/t/p/w154" & Trim(js.Item("results")(movienum).Item("poster_path"))
        Else

            get_TMDB_Poster = ""

        End If

    Else

        get_TMDB_Poster = ""

    End If
    
skip:
End Function
