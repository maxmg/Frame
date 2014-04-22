Imports Microsoft.VisualBasic
Imports System.Web
Imports System.Reflection
Imports System.Web.Script.Serialization
Imports System.Runtime.Serialization

Namespace CG.New

	' ============= Beginning of Classes ==================

    <Serializable()> _
    <DataContract()> _
    Public Class Project
        Friend Shared tableName As String = "projects"
        <DataMember()> Public id As New Field("id")
        <DataMember()> Public name As New Field("name")
        <DataMember()> Public typeLookup As New Field("project_type_lookup")
        <DataMember()> Public clubId As New Field("club_id")
        <DataMember()> Public shortDescription As New Field("description")
        <DataMember()> Public fullDescription As New Field("description")
        <DataMember()> Public peopleIds As New Field("people_ids")
        <DataMember()> Public studentId As New Field("student_id")
        <DataMember()> Public dateCreated As New Field("date")
        <DataMember()> Public updatedOn As New Field("updated_on")
        <DataMember()> Public targetDate As New Field("target_date", "date")
        <DataMember()> Public startDate As New Field("start_date", "date")
        <DataMember()> Public deleted As New Field("deleted")

        Public Sub New()
        End Sub

        Public Sub FormatData()
            updatedOn.val = CG.[New].Tools.writeShortDate(updatedOn.val)
        End Sub

    End Class 'Project

    <Serializable()> _
    <DataContract()> _
    Public Class Todo
        Friend Shared tableName As String = "to_dos"
        <DataMember()> Public id As New Field("id")
        <DataMember()> Public authorId As New Field("author_id")
        <DataMember()> Public title As New Field("title")
        <DataMember()> Public clubId As New Field("club_id")
        <DataMember()> Public description As New Field("description")
        <DataMember()> Public createdDate As New Field("created_date", "date")
        <DataMember()> Public startDate As New Field("start_date", "date")
        <DataMember()> Public targetDate As New Field("target_date", "date")
        <DataMember()> Public updatedAt As New Field("updated_at", "date")
        <DataMember()> Public projectId As New Field("project_id")
        <DataMember()> Public toPeopleId As New Field("to_people_id")
        <DataMember()> Public deleted As New Field("deleted")
        <DataMember()> Public statusLookup As New FieldLookup("status_lookup")
        <DataMember()> Public priorityLookup As New FieldLookup("priority_lookup")
        <DataMember()> Public peopleFirstName As New Field("first_name")
        <DataMember()> Public peopleLastName As New Field("last_name")
        <DataMember()> Public peoplePhotoUrl As New Field("to_people_id")

        Public Sub New()
        End Sub

        Public Sub FormatData()
            createdDate.val = CG.[New].Tools.writeShortDate(createdDate.val)
            startDate.val = CG.[New].Tools.writeShortDate(startDate.val)
            targetDate.val = CG.[New].Tools.writeShortDate(targetDate.val)
            updatedAt.val = CG.[New].Tools.writeShortDate(updatedAt.val)
            peopleFirstName.val = Left(peopleFirstName.val, 1)
            peopleLastName.val = Left(peopleLastName.val, 1)
            peoplePhotoUrl.val = CG.[New].Tools.getProfilePicture(peoplePhotoUrl.val, "thumb", CG.[New].Tools.getConnNet())
        End Sub

    End Class 'ToDo

    <Serializable()> _
    <DataContract()> _
    Public Class Note
        Friend Shared tableName As String = "notes"
        <DataMember()> Public id As New Field("id")
        <DataMember()> Public projectId As New Field("project_id")
        <DataMember()> Public title As New Field("title")
        <DataMember()> Public content As New Field("content")
        <DataMember()> Public shortContent As New Field("content")
        <DataMember()> Public createdDate As New Field("date", "date")
        <DataMember()> Public deleted As New Field("deleted")
        <DataMember()> Public updatedOn As New Field("updated_on")

        Public Sub New()
        End Sub

        Public Sub FormatData()
            shortContent.val = CG.[New].Tools.truncateText(shortContent.val, 25)
            createdDate.val = CG.[New].Tools.writeShortDate(createdDate.val)
        End Sub

    End Class 'Project

    <Serializable()> _
    <DataContract()> _
    Public Class WebPageTree
        Friend Shared tableName As String = "web_pages"
        <DataMember()> Public id As New Field("id")
        <DataMember()> Public uid As New Field("uid")
        <DataMember()> Public pageName As New Field("page_name")
        <DataMember()> Public clubId As New Field("club_id")
        <DataMember()> Public menuOrder As New Field("menu_order")
        <DataMember()> Public content As New Field("content")
        <DataMember()> Public subMenuOrder As New Field("sub_menu_order")
        <DataMember()> Public publish As New Field("publish")
        <DataMember()> Public homePage As New Field("home_page")
        <DataMember()> Public membersOnly As New Field("members_only")
        <DataMember()> Public type As New Field("type")
        <DataMember()> Public urlName As New Field("url_name")
        <DataMember()> Public deleted As New Field("deleted")
        <DataMember()> Public parentId As New Field("parent_id")

        Public Sub New()
        End Sub

        Public Sub FormatData()
            content.val = ""
        End Sub

    End Class 'WebPageTree

    <Serializable()> _
    <DataContract()> _
    Public Class Files
        Friend Shared tableName As String = "file_uploads"
        Public id As New Field("id")
        Public fileName As New Field("file_name")
        Public dateCreated As New Field("date")
        Public tags As New Field("tags")
        Public studentTags As New Field("student_tags")
        Public photoName As New Field("name")
        Public photoSize As New Field("size")
        Public deleted As New Field("deleted")

        Public Sub New()
        End Sub
    End Class 'Files

    <Serializable()> _
    <DataContract()> _
    Public Class Tag
        Friend Shared tableName As String = "tags"
        <DataMember()> Public id As New Field("id")
        <DataMember()> Public tagName As New Field("name")
        <DataMember()> Public className As New Field("class_name")
        <DataMember()> Public studentId As New Field("student_id")
        <DataMember()> Public type As New Field("type")
        <DataMember()> Public deleted As New Field("deleted")

        Public Sub New()
        End Sub
    End Class 'Tags

    ' ============= End of Classes ==================




    <Serializable()> _
    <DataContract()> _
    Partial Public Class Frame(Of T)

        <ScriptIgnore()> Public Query As String
        <ScriptIgnore()> Public Index As Integer
        Private myType As Type = GetType(T)
        Private myObject As T
        Private myFields As New List(Of String)
        Private myDBFields As New List(Of String)

        <DataMember()> Public obj As String = myType.Name
        <DataMember()> Public preAuth As String
        <DataMember()> Public items As New List(Of T)

        Public Sub Execute()

            loadFields()
            If Not String.IsNullOrEmpty(Query) Then
                Dim rsGetItems As Object
                Dim i As Integer
                CG.[New].Tools.openRec(rsGetItems, Query, CG.[New].Tools.getConnNet())

                If rsGetItems.eof Then
                    Dim item As T = Activator.CreateInstance(Of T)()
                    items.Add(item)
                End If

                Do While Not rsGetItems.eof
                    Dim item As T = Activator.CreateInstance(Of T)()
                    For i = 0 To myDBFields.Count() - 1
                        If Not String.IsNullOrEmpty(CG.[New].Tools.cleanObj(rsGetItems(myDBFields(i)).value)) Then
                            setField(item, myFields(i), "val", CG.[New].Tools.cleanObj(rsGetItems(myDBFields(i)).value))
                        End If
                    Next
                    For Each MethodObj In GetType(T).GetMethods()
                        If MethodObj.Name = "FormatData" Then
                            GetType(T).GetMethod("FormatData").Invoke(item, Nothing)
                        End If
                    Next
                    items.Add(item)
                    rsGetItems.moveNext()
                Loop

            End If

        End Sub

        Public Function loadFields() As List(Of String)
            getMyObject()
            Dim myMembersInfo As MemberInfo() = myType.GetFields()

            For i = 0 To myMembersInfo.Length - 1
                myFields.Add(myMembersInfo(i).Name)
                myDBFields.Add(getDBField(myMembersInfo(i).Name))
            Next
        End Function

        Public Function getDBField(ByVal fieldName As String) As String
            Dim myFieldInfo As FieldInfo = myType.GetField(fieldName)
            Dim myMemberType As Type = myFieldInfo.FieldType()
            Dim myFieldInfoMember As FieldInfo = myMemberType.GetField("dBField")
            Return myFieldInfoMember.GetValue(myFieldInfo.GetValue(myObject)).ToString()
        End Function

        Public Sub setField(ByVal obj As T, ByVal field As String, ByVal subfield As String, ByVal newVal As String)
            Dim myFieldInfo As FieldInfo = GetType(T).GetField(field)
            Dim myMemberType As Type = myFieldInfo.FieldType()
            Select Case (myMemberType.Name)
                Case "Field"
                    Dim myFieldInfoMember As FieldInfo = myMemberType.GetField(subfield)
                    myFieldInfoMember.SetValue(myFieldInfo.GetValue(obj), newVal)
                Case "FieldLookup"
                    Dim myFieldInfoMember As FieldInfo = myMemberType.GetField(subfield)
                    myFieldInfoMember.SetValue(myFieldInfo.GetValue(obj), newVal)
                    Dim strLabel As String = CG.[New].Tools.getLookup(newVal, CG.[New].Tools.getConnNet())
                    myMemberType.GetField("label").SetValue(myFieldInfo.GetValue(obj), strLabel)

                Case Else
                    'Calcul from the other fields
            End Select
        End Sub

        Public Sub getMyObject()
            myObject = GetType(T).InvokeMember(GetType(T).FullName, BindingFlags.DeclaredOnly Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.CreateInstance Or BindingFlags.NonPublic, Nothing, myObject, Nothing)
        End Sub

        Public Function Serialize(ByVal obj As Frame(Of T)) As String
            Dim jss As New System.Web.Script.Serialization.JavaScriptSerializer()
            Return jss.Serialize(obj)
        End Function

        Public Function Deserialize(ByVal json As String) As Frame(Of T)
            Dim obj = Activator.CreateInstance(Of Frame(Of T))()
            Using memoryStream As New System.IO.MemoryStream(Encoding.Unicode.GetBytes(json))
                Dim serializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(obj.GetType())
                memoryStream.Position = 0
                obj = serializer.ReadObject(memoryStream)
                Return obj
            End Using
        End Function

        Public Function saveObject(ByVal obj As List(Of T))
            getMyObject()
            Dim Response As String
            Dim jss As New System.Web.Script.Serialization.JavaScriptSerializer()
            Dim myMembersInfo As MemberInfo() = GetType(T).GetFields(BindingFlags.DeclaredOnly Or BindingFlags.Public Or BindingFlags.Instance)
            Dim rsUpdate As Object
            Dim Conn As Object = CG.[New].Tools.getConnNet()
            'Try
            For k = 0 To obj.Count - 1
                For i = 0 To myMembersInfo.Length - 1
                    Dim myFieldInfo As FieldInfo = GetType(T).GetField(myMembersInfo(i).Name)
                    Dim myMemberType As Type = myFieldInfo.FieldType
                    Dim myFieldInfoMember As FieldInfo = myMemberType.GetField("u")
                    If myFieldInfoMember.GetValue(myFieldInfo.GetValue(obj(k))) = 1 Then
                        myFieldInfoMember = myMemberType.GetField("val")
                        Dim myFieldId As Object = GetType(T).GetField("id").FieldType.GetField("val").GetValue(GetType(T).GetField("id").GetValue(obj(k)))
                        Dim myDBField As Object = GetType(T).GetField(myMembersInfo(i).Name).FieldType.GetField("dBField").GetValue(GetType(T).GetField(myMembersInfo(i).Name).GetValue(myObject))
                        Dim myTypeField As Object = GetType(T).GetField(myMembersInfo(i).Name).FieldType.GetField("type").GetValue(GetType(T).GetField(myMembersInfo(i).Name).GetValue(myObject))

                        Dim strValue As String = myMemberType.GetField("val").GetValue(myFieldInfo.GetValue(obj(k)))

                        If IsNumeric(myFieldId) Then
                            If myFieldId > 0 Then
                                If myTypeField = "date" Then
                                    If Not String.IsNullOrEmpty(strValue) Then
                                        strValue = CG.[New].Tools.mySqlDateTime(strValue)
                                    Else
                                        strValue = ""
                                    End If
                                Else
                                    strValue = CG.[New].Tools.stripReplace(strValue)
                                End If
                                If String.IsNullOrEmpty(strValue) Then
                                    strValue = "NULL"
                                Else
                                    strValue = "'" & strValue & "'"
                                End If
                                HttpContext.Current.Response.Write("UPDATE " & getTableName() & " SET " & myDBField & " = " & strValue & " WHERE id = " & myFieldId)
                                CG.[New].Tools.openRec(rsUpdate, "UPDATE " & getTableName() & " SET " & myDBField & " = " & strValue & " WHERE id = " & myFieldId, Conn)
                            ElseIf myFieldId = -1 Then
                                CG.[New].Tools.openRec(rsUpdate, "INSERT INTO " & getTableName() & " (sessionid) VALUES ('" & HttpContext.Current.Session.SessionID & "')", Conn)
                                setField(obj(k), "id", "val", CG.[New].Tools.getLastNew("id", getTableName(), "id", HttpContext.Current.Session.SessionID, Conn))
                                i = i - 1
                            End If
                        End If
                    End If
                Next
            Next
            Response = "{""success"":true}"
            'Catch ex As Exception
            '   CG.[New].Tools.openRec(rsUpdate, "INSERT INTO errors (sessionid, type, date, student_id, url, message) VALUES ('" & HttpContext.Current.Session.SessionID & "', 'ASPX-DATAMODEL', '" & CG.[New].Tools.mySqlDateTime(Date.Now) & "', '" & HttpContext.Current.Request.QueryString.ToString() & "', '" & HttpContext.Current.Session("student_id") & "', '" & CG.[New].Tools.stripReplace(ex.Message) & "')", Conn)
            '  Response = "false"
            'End Try
            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.Write(Response)
            HttpContext.Current.Response.End()
        End Function

        Public Function getTableName()
            Dim sharedField As FieldInfo = GetType(T).GetField("tableName", BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static)
            Return sharedField.GetValue(myObject)
        End Function

        Public Sub New()
        End Sub

    End Class 'Frame(Of T)

    <Serializable()> _
    <DataContract()> _
    Public Class Field
        <DataMember()> Public val As String
        <DataMember()> Public u As Integer
        <ScriptIgnore()> Public dBField As String
        <ScriptIgnore()> Public type As String

        Public Sub New(ByVal dbFieldName As String)
            dBField = dbFieldName
        End Sub

        Public Sub New(ByVal dbFieldName As String, ByVal dbType As String)
            dBField = dbFieldName
            type = dbType
        End Sub

    End Class 'Field

    <Serializable()> _
    <DataContract()> _
    Public Class FieldTable
        <DataMember()> Public val As String
        <DataMember()> Public u As Integer
        <ScriptIgnore()> Public dBField As String
        <ScriptIgnore()> Public dBTable As String

        Public Sub New(ByVal dbFieldName As String, ByVal dbTableName As String)
            dBField = dbFieldName
            dBTable = dbTableName
        End Sub

    End Class 'Field

    <Serializable()> _
    <DataContract()> _
    Public Class FieldLookup
        <DataMember()> Public val As String
        <DataMember()> Public u As Integer
        <DataMember()> Public label As String
        <ScriptIgnore()> Public dBField As String
        <ScriptIgnore()> Public type As String

        Public Sub New(ByVal dbFieldName As String)
            dBField = dbFieldName
        End Sub

    End Class 'FieldLookup

    Public Class FieldFunction
        <DataMember()> Public val As String
        Public Sub New()
        End Sub
    End Class

End Namespace
