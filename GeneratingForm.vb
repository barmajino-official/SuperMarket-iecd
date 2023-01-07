Imports Google.Cloud.Firestore ' This line imports the Google Cloud Firestore library, which is used to access a Cloud Firestore database
Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows.Forms ' This line imports the System.Windows.Forms library, which contains classes for creating Windows-based applications

Module GeneratingForm ' This line declares a module (a container for procedures and variables) called "AddProduct"
    Dim form As New Form() ' This line declares a Form object called "form" and instantiates it using the Form class's default constructor
    Dim productsDataGridView As New DataGridView() ' This line declares a DataGridView object called "productsDataGridView" and instantiates it using the DataGridView class's default constructor
    Dim productsData As New DataTable()
    Public db As FirestoreDb
    Public menuStrip As MenuStrip
    Dim RecordSelected As String
    Dim TheDialog As New Form()
    Public ItemS As String
    Dim columnHeaders As New List(Of Tuple(Of String, String)) ' This line declares a List object called "columnHeaders" and initializes it with eight string values


    Public Function GeneratingForm(main As Form) ' This line declares a public function called "AddProductsForm" that returns a Form object
        ' Set the form properties
        With form
            .MdiParent = main
            .Text = ItemS ' This line sets the Text property of the "form" object to the string "Products"
            .ControlBox = False ' This line sets the ControlBox property of the "form" object to False
            .ShowIcon = False ' This line sets the ShowIcon property of the "form" object to False
            .FormBorderStyle = FormBorderStyle.None
            .Dock = DockStyle.Fill
            .WindowState = FormWindowState.Normal
        End With




        'AddHandler form.Load, AddressOf LoadForm ' This line registers an event handler for the Load event of the "form" object, which will call the "LoadForm" subroutine when the event is raised
        ' Show the form
        Return form ' This line returns the "form" object
    End Function


    Public Sub SetForm() ' This line declares a private subroutine called "LoadForm" that is called when the form is loaded
        menuStrip.Items.Add("Add", Nothing, Sub() OpenAddDialog())
        menuStrip.Items.Add("Edit", Nothing, Sub() OpenEditDialog())
        menuStrip.Items.Add("Delete", Nothing, Sub() DeleteProductFromDb())
        menuStrip.Items.Add("Enable Filter", Nothing, Sub() E_D_Filter())

        ' ///////////////////////////////////////////
        ' First, create a new ToolStripTextBox
        Dim textBox As New ToolStripTextBox()

        ' Set the textbox properties
        textBox.Name = "Filter Table"
        ' Add the menu item to the menu strip
        menuStrip.Items.Add(textBox)
        AddHandler menuStrip.Items(menuStrip.Items.IndexOf(textBox)).TextChanged, AddressOf Filterdata


        '///////////////////////////////////////////////
        menuStrip.Items.Add("Clear", Nothing, Sub() menuStrip.Items(4).Text = "")
        menuStrip.Items(4).Available = False
        menuStrip.Items(5).Available = False

        columnHeaders.Add(New Tuple(Of String, String)("Id", $"{ItemS} ID"))
        For Each data As Object In GetConfig(ItemS)
            columnHeaders.Add(New Tuple(Of String, String)(data("Title"), data("Key")))
        Next


        ' Create a new DataGridView control and add it to the form
        With productsDataGridView ' This line begins a block of code that operates on the "productsDataGridView" object
            .Name = "Product" ' This line sets the Name property of the "productsDataGridView" object to the string "Product"
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .AllowUserToAddRows = False
            .AllowUserToResizeRows = False
            .EditMode = False
            .Width = form.ClientSize.Width
            .Location = New Point(0, 0)
            .Dock = DockStyle.Fill
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        End With
        AddHandler productsDataGridView.CellClick, AddressOf GetIdOfSelectedRow
        productsDataGridView.ColumnHeadersVisible = True
        For i = 0 To columnHeaders.Count - 1
            productsData.Columns.Add(columnHeaders(i).Item1, GetType(String))
            productsData.Columns(i).ReadOnly = True
        Next
        form.Controls.Add(productsDataGridView)
        ' Create a new menuStrip object

        Call GetAllRecords()
    End Sub

    Private Sub E_D_Filter()
        If menuStrip.Items(3).Text = "Enable Filter" Then
            menuStrip.Items(3).Text = "Disable Filter"
            menuStrip.Items(4).Available = True
            menuStrip.Items(5).Available = True
            Filterdata()
        Else
            menuStrip.Items(3).Text = "Enable Filter"
            menuStrip.Items(4).Available = False
            menuStrip.Items(5).Available = False
            Dim dataSourceType As Type = productsDataGridView.DataSource.GetType()
            If TypeOf productsDataGridView.DataSource Is DataTable Then
                ' Data source is a DataTable, so you can apply the filter
                TryCast(productsDataGridView.DataSource, DataTable).DefaultView.RowFilter = ""
            Else
                ' Data source is not a DataTable, so you cannot apply the filter
                MessageBox.Show("Cannot filter data source. Data source is not a DataTable.")
            End If
        End If

    End Sub
    Private Sub Filterdata()

        ' Get the text to filter by
        Dim filterText As String = menuStrip.Items(4).Text

        ' Build the filter expression
        Dim filter As String = ""
        For Each column As DataGridViewColumn In productsDataGridView.Columns
            If filter <> "" Then
                filter &= " OR "
            End If
            filter &= $"[{column.Name}] LIKE '%{filterText}%'"
        Next

        Dim dataSourceType As Type = productsDataGridView.DataSource.GetType()
        If TypeOf productsDataGridView.DataSource Is DataTable Then
            ' Data source is a DataTable, so you can apply the filter
            TryCast(productsDataGridView.DataSource, DataTable).DefaultView.RowFilter = String.Format(filter)
        Else
            ' Data source is not a DataTable, so you cannot apply the filter
            MessageBox.Show("Cannot filter data source. Data source is not a DataTable.")
        End If



        ' Set the DataGridView's data source to the filtered view
    End Sub
    ' **********************************************************************************************

    Private Sub OpenAddDialog()
        ' Create a new form
        TheDialog.ControlBox = False
        ' Set the text of the form
        TheDialog.Text = "Add Items"

        ' Set the size of the form



        Dim y_slide As Integer = 10
        Dim x_slide As Integer = (300 / 2) - 100
        Dim thedata As List(Of Object) = GetConfig(ItemS)
        For Each data_ As Dictionary(Of String, Object) In thedata
            Dim lable As Label = New Label()
            With lable
                .Text = data_("Title")
                .Width = 100
                .Location = New Point(x_slide, y_slide)
                .TextAlign = ContentAlignment.MiddleCenter
            End With
            Dim control As Control = GetSpectypes(data_("Control"))
            If TypeOf control Is ComboBox Then
                Dim comboBox As ComboBox = CType(control, ComboBox)
                comboBox.Items.AddRange(data_("Data").ToArray())
                comboBox.SelectedIndex = 0
            End If
            If TypeOf control Is NumericUpDown Then
                Dim NumericUpDown_ As NumericUpDown = CType(control, NumericUpDown)
                With NumericUpDown_
                    .Minimum = data_("Config")("Min")
                    .Maximum = data_("Config")("Max")
                    .Value = data_("Config")("Curent")
                    .Visible = data_("Config")("Visible")("Add")
                End With
            End If
            If TypeOf control Is Label Then
                Dim labelBox As Label = CType(control, Label)
                labelBox.Text = data_("Config")("Curent")
                labelBox.TextAlign = ContentAlignment.MiddleCenter
            End If
            If TypeOf control Is TextBox Then
                Dim labelBox As TextBox = CType(control, TextBox)
                labelBox.Text = data_("Config")("Curent")
            End If
            If data_("Config")("Visible")("Add") Then
                control.Width = 120
                control.Location = New Point(x_slide + 100, y_slide)
                TheDialog.Controls.Add(control)
                TheDialog.Controls.Add(lable)
                y_slide += 30
            End If
            DataControls.Add(data_("Key"), control)
        Next

        ' Create a new button
        Dim B_Add As New Button()
        ' Set the text of the button
        B_Add.Text = "Add"
        B_Add.Width = 100
        B_Add.Location = New Point(x_slide, y_slide)
        ' Handle the click event of the button
        AddHandler B_Add.Click, AddressOf SaveValues
        ' Add the button to the form
        TheDialog.Controls.Add(B_Add)

        ' Create a new button
        Dim B_cancel As New Button()
        ' Set the text of the button
        B_cancel.Text = "Cancel"
        B_cancel.Width = 120
        B_cancel.Location = New Point(x_slide + 100, y_slide)
        ' Handle the click event of the button
        AddHandler B_cancel.Click, AddressOf CloseTheDialog
        ' Add the button to the form
        TheDialog.Controls.Add(B_cancel)

        TheDialog.FormBorderStyle = FormBorderStyle.FixedDialog
        TheDialog.MaximizeBox = False
        TheDialog.Size = New Size(x_slide + 300, y_slide + 70)
        ' Display the form as a modal dialog box
        TheDialog.ShowDialog()
    End Sub

    Private Sub CloseTheDialog()
        TheDialog.Close()
        TheDialog.Controls.Clear()
        DataControls.Clear()
        RecordSelected = ""
        productsDataGridView.ClearSelection()
    End Sub

    Private Sub OpenEditDialog()
        If RecordSelected = "" Then
            MessageBox.Show($"Please select {ItemS} first then click the Action")

        Else
            Dim Permition As Integer = MessageBox.Show($"Are you sur you want to Edit record's ID : {RecordSelected} ?", "Edit Alert", MessageBoxButtons.YesNo)
            If Permition = 6 Then
                Dim productDocRef As DocumentReference = db.Collection(ItemS).Document(RecordSelected)
                Dim productDoc As DocumentSnapshot = productDocRef.GetSnapshotAsync().Result

                Dim EditRecordData As Dictionary(Of String, Object) = productDoc.ToDictionary()

                ' Create a new form
                TheDialog.ControlBox = False
                ' Set the text of the form
                TheDialog.Text = "Edit Items"

                ' Set the size of the form



                Dim y_slide As Integer = 10
                Dim x_slide As Integer = (300 / 2) - 100
                Dim thedata As List(Of Object) = GetConfig(ItemS)
                Dim EditRecordvalue As Object
                For Each data_ As Dictionary(Of String, Object) In thedata

                    If EditRecordData.ContainsKey(data_("Key")) Then

                        EditRecordvalue = EditRecordData(data_("Key"))
                    Else
                        EditRecordvalue = "Null"
                    End If



                    Dim lable As Label = New Label()
                    With lable
                        .Text = data_("Title")
                        .Width = 100
                        .Location = New Point(x_slide, y_slide)
                        .TextAlign = ContentAlignment.MiddleCenter
                    End With
                    Dim control As Control = GetSpectypes(data_("Control"))
                    If TypeOf control Is ComboBox Then
                        Dim comboBox As ComboBox = CType(control, ComboBox)
                        comboBox.Items.AddRange(data_("Data").ToArray())
                        comboBox.SelectedIndex = comboBox.Items.IndexOf(EditRecordvalue)

                    End If
                    If TypeOf control Is NumericUpDown Then
                        Dim NumericUpDown_ As NumericUpDown = CType(control, NumericUpDown)
                        With NumericUpDown_
                            .Minimum = data_("Config")("Min")
                            .Maximum = data_("Config")("Max")
                            .Value = If(EditRecordvalue Is "Null", 0, EditRecordvalue)
                            .Visible = data_("Config")("Visible")("Edit")
                        End With
                    End If
                    If TypeOf control Is Label Then
                        Dim labelBox As Label = CType(control, Label)
                        labelBox.Text = EditRecordvalue
                        labelBox.TextAlign = ContentAlignment.MiddleCenter
                    End If
                    If TypeOf control Is TextBox Then
                        Dim labelBox As TextBox = CType(control, TextBox)
                        labelBox.Text = EditRecordvalue
                    End If
                    If data_("Config")("Visible")("Edit") Then
                        control.Width = 120
                        control.Location = New Point(x_slide + 100, y_slide)
                        TheDialog.Controls.Add(control)
                        TheDialog.Controls.Add(lable)
                        y_slide += 30
                    End If
                    DataControls.Add(data_("Key"), control)
                Next

                ' Create a new button
                Dim B_Add As New Button()
                ' Set the text of the button
                B_Add.Text = "Update"
                B_Add.Width = 100
                B_Add.Location = New Point(x_slide, y_slide)
                ' Handle the click event of the button
                AddHandler B_Add.Click, AddressOf EditValues
                ' Add the button to the form
                TheDialog.Controls.Add(B_Add)

                ' Create a new button
                Dim B_cancel As New Button()
                ' Set the text of the button
                B_cancel.Text = "Cancel"
                B_cancel.Width = 120
                B_cancel.Location = New Point(x_slide + 100, y_slide)
                ' Handle the click event of the button
                AddHandler B_cancel.Click, AddressOf CloseTheDialog
                ' Add the button to the form
                TheDialog.Controls.Add(B_cancel)

                TheDialog.FormBorderStyle = FormBorderStyle.FixedDialog
                TheDialog.MaximizeBox = False
                TheDialog.Size = New Size(x_slide + 300, y_slide + 70)
                ' Display the form as a modal dialog box
                TheDialog.ShowDialog()


                Call GetAllRecords()
            Else
                CloseTheDialog()
            End If
        End If
    End Sub
    ' **********************************************************************************************

    Public Sub GetIdOfSelectedRow() ' This line declares a public subroutine called "GetIdOfSelectedRow" that is called when a cell in the "productsDataGridView" object is clicked
        ' Get a reference to the row that was clicked
        Dim rowIndex As Integer = productsDataGridView.CurrentCell.RowIndex ' This line declares a variable called "rowIndex" and assigns it the value of the row index of the currently selected cell in the "productsDataGridView" object

        Dim row As DataGridViewRow = productsDataGridView.Rows(rowIndex) ' This line declares a variable called "row" and assigns it the value of the row at the index specified by the "rowIndex" variable

        ' Your code to handle the click event here
        RecordSelected = row.Cells(0).Value
    End Sub


    Public Sub ClearForm()
        productsDataGridView.DataSource = Nothing
        columnHeaders.Clear()
        form.Controls.Clear()

        productsData.Columns.Clear()
        productsData.Rows.Clear()
    End Sub




    Private Sub GetAllRecords()
        Try
            Dim productsCollection As CollectionReference = db.Collection(ItemS)
            Dim querySnapshot As QuerySnapshot = productsCollection.GetSnapshotAsync().Result
            productsData.Rows.Clear()
            For Each document As Object In querySnapshot.Documents
                Dim data As Dictionary(Of String, Object) = document.ToDictionary()
                data.Add($"{ItemS} ID", document.Id)
                Dim row1 As List(Of String) = New List(Of String)
                data = ReorderData(columnHeaders, data)
                For Each pair As KeyValuePair(Of String, Object) In data
                    Dim value_ As String = pair.Value.ToString()
                    row1.Add(value_)
                Next
                productsData.Rows.Add(row1.ToArray())

            Next
            productsDataGridView.DataSource = productsData
        Catch ex As Exception
            MessageBox.Show($"Error retrieving documents: {ex.Message}")
        End Try
        productsDataGridView.ClearSelection()
    End Sub



    Private Sub DeleteProductFromDb()
        If RecordSelected = "" Then
            MessageBox.Show($"Please select {ItemS} first then click the Action")
        Else
            Dim Permition As Integer = MessageBox.Show($"Are you sur you want to delete record's ID : {RecordSelected} ?", "Delete", MessageBoxButtons.YesNo)
            If Permition = 6 Then
                Dim docRef As DocumentReference = db.Collection(ItemS).Document(RecordSelected)
                docRef.DeleteAsync().Wait()
                RecordSelected = ""

                Call GetAllRecords()

            End If
        End If
    End Sub

    Public Sub SaveValues()
        Dim data As New Dictionary(Of String, Object)
        For Each dataControl As KeyValuePair(Of String, Control) In DataControls
            Dim control As Control = dataControl.Value
            Dim controlType As Type = control.GetType()
            Dim value As Object = Nothing
            For Each member As MemberInfo In controlType.GetMembers()
                If member.MemberType = MemberTypes.Property OrElse member.MemberType = MemberTypes.Method Then
                    Dim method As MethodInfo = TryCast(member, MethodInfo)
                    If method IsNot Nothing AndAlso method.ReturnType <> GetType(Void) Then
                        value = method.Invoke(control, Nothing)
                        Exit For
                    Else
                    End If
                End If
            Next
            value = If(value = 128, control.Text, value)
            If value.GetType() = GetType(Decimal) Then
                value = CDbl(value)
            End If
            If TypeOf control Is TextBox Then
                Dim textBox As TextBox = CType(control, TextBox)
                value = textBox.Text
            End If
            If dataControl.Key = "SysCD" Then
                Dim currentTime As DateTime = DateTime.Now
                Dim formattedTime As String = currentTime.ToString("MM/dd/yyyy hh:mm:ss tt")
                data.Add(dataControl.Key, formattedTime)
            Else
                data.Add(dataControl.Key, value)
            End If

        Next
        Dim usersCollection As CollectionReference = db.Collection(ItemS)

        Dim doc As DocumentReference = usersCollection.AddAsync(data).Result
        MessageBox.Show($"Added document with ID: {doc.Id}")
        CloseTheDialog()

        Call GetAllRecords()
    End Sub

    Public Sub EditValues()
        Dim data As New Dictionary(Of String, Object)
        For Each dataControl As KeyValuePair(Of String, Control) In DataControls
            Dim control As Control = dataControl.Value
            Dim controlType As Type = control.GetType()
            Dim value As Object = Nothing
            For Each member As MemberInfo In controlType.GetMembers()
                If member.MemberType = MemberTypes.Property OrElse member.MemberType = MemberTypes.Method Then
                    Dim method As MethodInfo = TryCast(member, MethodInfo)
                    If method IsNot Nothing AndAlso method.ReturnType <> GetType(Void) Then
                        value = method.Invoke(control, Nothing)
                        Exit For
                    Else
                    End If
                End If
            Next
            value = If(value = 128, control.Text, value)
            If value.GetType() = GetType(Decimal) Then
                value = CDbl(value)
            End If
            If TypeOf control Is TextBox Then
                Dim textBox As TextBox = CType(control, TextBox)
                value = textBox.Text
            End If
            If dataControl.Key <> "SysCD" Then
                data.Add(dataControl.Key, value)
            End If

        Next
        Dim productDocRef As DocumentReference = db.Collection(ItemS).Document(RecordSelected)

        productDocRef.UpdateAsync(data)

        MessageBox.Show($"Updated  document with ID: {RecordSelected}")
        CloseTheDialog()

        Call GetAllRecords()
    End Sub

End Module
