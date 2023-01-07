Imports System.Reflection
Imports System.Windows.Forms
Imports Google.Cloud.Firestore

Module BuildComponents
    Public db As FirestoreDb
    Public DataControls As Dictionary(Of String, Control) = New Dictionary(Of String, Control)
    Public Function GetSpectypes(name As String) As Control
        For Each type As Type In GetType(System.Windows.Forms.Control).Assembly.GetTypes()
            If type.IsSubclassOf(GetType(System.Windows.Forms.Control)) Then
                If type.Name.ToString().ToLower() = name.ToLower() Then
                    Dim control As Control = DirectCast(Activator.CreateInstance(type), Control)
                    Return control
                End If
            End If
        Next

        Return Nothing
    End Function

    Function CheckVariable(ByVal variable As Object) As Boolean
        If IsNothing(variable) Then
            Return False
        ElseIf TypeOf variable Is String Then
            Return True
        ElseIf TypeOf variable Is Boolean Then
            Return variable
        Else
            Return False
        End If
    End Function

    Public Function ReorderData(columnOrder As List(Of Tuple(Of String, String)), inputData As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
        Dim reorderedData As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
        For Each columnName As Tuple(Of String, String) In columnOrder
            If inputData.ContainsKey(columnName.Item2) Then
                reorderedData.Add(columnName.Item2, inputData(columnName.Item2))
            Else
                reorderedData.Add(columnName.Item2, "Null")
            End If
        Next
        Return reorderedData
    End Function

    Public Function GetConfig(configOf As String) As List(Of Object)
        Dim db As FirestoreDb = FirestoreDb.Create("supermarket-5d271")

        Dim controlsObjects As Dictionary(Of String, Control) = New Dictionary(Of String, Control)
        Dim productDocRef As DocumentReference = db.Collection("Configurations").Document(configOf)
        Dim productDoc As DocumentSnapshot = productDocRef.GetSnapshotAsync().Result
        Dim data As Dictionary(Of String, Object) = productDoc.ToDictionary()
        For Each pair As KeyValuePair(Of String, Object) In data
            Dim theData As List(Of Object) = pair.Value
            Return theData
        Next

        Return Nothing
    End Function

    Public Function GetMenue() As List(Of String)
        Dim Listname As New List(Of String)
        Dim productsCollection As CollectionReference = db.Collection("Configurations")
        Dim querySnapshot As QuerySnapshot = productsCollection.GetSnapshotAsync().Result
        For Each document As Object In querySnapshot.Documents
            Dim data As Dictionary(Of String, Object) = document.ToDictionary()
            Listname.Add(document.Id)
        Next
        Return Listname
    End Function
End Module
