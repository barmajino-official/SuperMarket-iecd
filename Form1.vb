Imports Google.Cloud.Firestore
Imports System.Windows.Forms

Public Class Main
    Dim Data_ As Dictionary(Of String, TextBox) = New Dictionary(Of String, TextBox)()
    Dim dialog As New Form()
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim formm As Form = Me
        'Me.IsMdiContainer = True
        formm.IsMdiContainer = True

        Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", "src/supermarket-5d271-firebase-adminsdk-8l7cy-f096d393a7.json")
        ' Create a FirestoreDb object to interact with Cloud Firestore.
        Dim db As FirestoreDb = FirestoreDb.Create("supermarket-5d271")
        GeneratingForm.db = db
        BuildComponents.db = db
        Dim menuStripTop As MenuStrip = New MenuStrip()
        menuStripTop.Dock = DockStyle.Top

        For Each name As String In GetMenue()

            menuStripTop.Items.Add(name, Nothing, Sub() openapp_(name))
        Next

        'menuStripTop.Items.Add("hi", Nothing, Sub() MessageBox.Show("hi"))

        formm.Controls.Add(menuStripTop)


        menuStrip = New MenuStrip()
        menuStrip.Dock = DockStyle.Bottom

        Dim theDynamicForm As Form = GeneratingForm.GeneratingForm(formm)
        formm.Controls.Add(menuStrip)
        theDynamicForm.Show()

        formm.Show()
    End Sub

    Private Sub openapp_(item As String)
        ClearForm()
        menuStrip.Items.Clear()
        ItemS = item
        SetForm()
    End Sub




End Class
