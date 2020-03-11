Imports System.Data.Odbc
Imports System.Data.OleDb

Public Class Principale
    Public myConnectionStringGalileo As String = "Driver={Client Access ODBC Driver (32-bit)};" & "System=10.1.128.250;TRANSLATE=1;Uid=ODBC;Pwd=ODBC"
    Public myConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\server\MAGIC PACK\PROCEDURE\DATABASE\Documenti_Fatturazione.mdb;"
    'Public myConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\server\MAGIC PACK\PROCEDURE\DATABASE\Documenti_Fatturazione_DEBUG.mdb;"
    Public odbcGalileo As New OdbcConnection
    Public myDA As New OdbcDataAdapter
    Public tabDDT As New DataTable
    Public rowDDT As DataRow
    Public tabBOL As New DataTable
    Public rowBOL As DataRow
    Public tabCLI As New DataTable
    Public rowCLI As DataRow
    Public tabIVA As New DataTable
    Public rowIVA As DataRow
    Public tabDocumentiArchivio As New DataTable
    Public rowDocumentiArchivio As DataRow
    Public tabAnag1503 As New DataTable
    Public rowAnag1503 As DataRow
    Public tabVettori As New DataTable
    Public rowVettori As DataRow
    Public tabCheckMovimenti As New DataTable
    Public rowCheckMovimenti As DataRow
    Public tabStampa As New DataTable
    Public rowStampa As DataRow
    Public TMPtabBol As New DataTable ' TEMPORANEA PER L'INSERIMENTO DI MASSA DEI DOCUMENTI VIA N° BOLLA
    Public TMProwBol As DataRow
    Public query As String = ""
    Public Mese As String = ""
    Public MeseDopo As String = ""
    Public AnnoDopo As String = ""
    Public Anno As String = ""
    Public DataI As String = ""
    Public DataF As String = ""
    Public DataInizioAnno As String = ""
    Public RagsPrimoVet As String = ""
    Public RagsSecondoVet As String = ""
    Public contaDocPresenti As Integer = 0
    Public contaDocMancanti As Integer = 0
    Public maxRiga As Integer = 0
    Public flag_vettore As String
    Public flag_cliente As String
    Public cnDB As New OleDbConnection
    Public myDAccess As New OleDbDataAdapter
    Public myCommand As New OleDbCommand
    Public srcViaNumDoc As Boolean = True
    Public changeSrcMode As Boolean = False
    Public Happy As String = "1501000603"
    Public Bollettario As String = " "
    Public Stato As String = " "
    Public lastNrDoc As String = ""

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim ianno As Integer
        Dim ii As Integer

        ianno = Today.Year
        For ii = ianno To ianno - 10 Step -1
            CmbAnno.Items.Add(ii.ToString)
            CmbAnno1.Items.Add(ii.ToString)
            CmbAnnoDoc.Items.Add(ii.ToString)
        Next

        CmbAnno.SelectedIndex = 0
        CmbAnno1.SelectedIndex = 0
        CmbMese.SelectedIndex = (Today.Month - 1)
        CmbMese1.SelectedIndex = (Today.Month - 1)

        query = "SELECT trim(MAG80DAT.SMTAB00F.XCODTB) AS CODICE_VETTORE, " & _
                " trim(substring(XDATTB,29,40)) AS VETTORE " & _
                " FROM MAG80DAT.SMTAB00F" & _
                " WHERE (((MAG80DAT.SMTAB00F.XTIPTB)='01TR'))"

        odbcGalileo.ConnectionString = myConnectionStringGalileo
        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabVettori.Clear()
        myDA.Fill(tabVettori)
        odbcGalileo.Close()

        DatA.Value = Today.AddDays(-1)
        DatDA.Value = Today.AddDays(-1)
        DtpDa369.Value = Today.AddDays(-1)
        DtpA369.Value = Today.AddDays(-1)

        dtpStampa.Value = Today.AddDays(-1)

        LblAnnoDoc.Visible = False
        LblBollettarioDoc.Visible = False
        CmbBollettarioDoc.Visible = False
        CmbAnnoDoc.Visible = False
    End Sub

    Private Sub BtnControlla1_Click(sender As System.Object, e As System.EventArgs) Handles BtnControlla1.Click
        Dim NrDDTprecedente As Integer = 0
        Dim DataPrima As String = ""
        Dim ClienteP As String = ""
        Dim PagamentoP As String = ""
        Dim AgenteP As String = ""
        Dim NomeAgenteP As String = ""
        Dim NomeClienteP As String = ""
        Dim AbiP As String = ""
        Dim CabP As String = ""
        Dim NomePagamentoP As String = ""
        Dim QtaP As Integer = 0
        Dim Diversi As Integer = 0
        Dim Cliente As String = ""
        Dim Pagamento As String = ""
        Dim Agente As String = ""
        Dim Abi As String = ""
        Dim Cab As String = ""

        Dim tmpTabControlli As New DataTable
        Dim tmpTabAgenti As New DataTable
        Dim tmpTabPagamenti As New DataTable

        Cursor = Cursors.WaitCursor

        Mese = CmbMese1.SelectedItem
        Anno = CmbAnno1.SelectedItem
        MeseDopo = Mese + 1
        AnnoDopo = Anno
        If MeseDopo = "13" Then
            MeseDopo = "01"
            AnnoDopo = Anno + 1
        ElseIf CInt(MeseDopo) < 10 Then
            MeseDopo = "0" & MeseDopo
        End If

        DataI = Anno & Mese & "01"
        DataF = AnnoDopo & MeseDopo & "01"
        DataInizioAnno = Anno & "01" & "01"

        DgvDati1.Rows.Clear()

        'query = "select MAG80DAT.FTMOV00F.CDCFFM AS CLIENTE, " & _
        '        "MAG80DAT.CGANA01J.DSCOCP AS RAGSOC, " & _
        '        "MAG80DAT.FTMOV00F.CDPAFM as PAGAMENTO, " & _
        '        "MAG80DAT.FTMOV00F.CDAGFM AS AGENTE, " & _
        '        "rtrim(SUBSTRING(MAG80DAT.SMTAB00F.XDATTB,9,20)) AS NOME_AGENTE, " & _
        '        "rtrim(SUBSTRING(PAGA.XDATTB,9,40)) AS DESCR_PAGAMENTO, " & _
        '        "count(*) as QTA " & _
        '        "from ((MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTMOV00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA) " & _
        '        "LEFT JOIN MAG80DAT.SMTAB00F ON MAG80DAT.FTMOV00F.CDAGFM = MAG80DAT.SMTAB00F.XCODTB) " & _
        '        "LEFT JOIN MAG80DAT.SMTAB00F PAGA ON MAG80DAT.FTMOV00F.CDPAFM = PAGA.XCODTB " & _
        '        "where MAG80DAT.FTMOV00F.DTBOFM >= " & DataI & " and MAG80DAT.FTMOV00F.DTBOFM <= " & DataF & " " & _
        '        "and MAG80DAT.FTMOV00F.TDOCFM='B' and MAG80DAT.FTMOV00F.CDCFFM <> '" & Happy & "' " & _
        '        "and MAG80DAT.SMTAB00F.XTIPTB Like '01AG%' and PAGA.XTIPTB='01CP' " & _
        '        "GROUP BY MAG80DAT.FTMOV00F.CDCFFM, MAG80DAT.CGANA01J.DSCOCP, MAG80DAT.FTMOV00F.CDPAFM, " & _
        '        "rtrim(SUBSTRING(MAG80DAT.SMTAB00F.XDATTB,9,20)), " & _
        '        "MAG80DAT.FTMOV00F.CDAGFM,rtrim(SUBSTRING(PAGA.XDATTB,9,40)) " & _
        '        "order by MAG80DAT.FTMOV00F.CDCFFM, MAG80DAT.FTMOV00F.CDPAFM, MAG80DAT.FTMOV00F.CDAGFM "

        query = " select MAG80DAT.SMTAB00F.XCODTB AS CODICE_AGENTE," & _
                " rtrim(SUBSTRING(MAG80DAT.SMTAB00F.XDATTB,9,20)) AS DESCRIZIONE_AGENTE" & _
                " from MAG80DAT.SMTAB00F " & _
                " WHERE  MAG80DAT.SMTAB00F.XTIPTB Like '01AG%'"

        odbcGalileo.ConnectionString = myConnectionStringGalileo

        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tmpTabAgenti.Clear()
        myDA.Fill(tmpTabAgenti)

        query = " select MAG80DAT.SMTAB00F.XCODTB AS CODICE_PAGAMENTO," & _
                " rtrim(SUBSTRING(MAG80DAT.SMTAB00F.XDATTB,9,20)) AS DESCRIZIONE_PAGAMENTO" & _
                " from MAG80DAT.SMTAB00F " & _
                " WHERE  MAG80DAT.SMTAB00F.XTIPTB Like '01CP%'"

        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tmpTabPagamenti.Clear()
        myDA.Fill(tmpTabPagamenti)

        query = "select MAG80DAT.FTMOV00F.CDCFFM AS CLIENTE, " & _
                "MAG80DAT.CGANA01J.CDPGCA AS CODICE_PAGAMENTO," & _
                "MAG80DAT.CGANA01J.CDAGCA AS COD_AGENTE," & _
                "MAG80DAT.CGANA01J.CABICA AS COD_ABI, " & _
                "MAG80DAT.CGANA01J.CCABCA AS COD_CAB " & _
                "from (MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTMOV00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA) " & _
                "where MAG80DAT.FTMOV00F.DTBOFM >= " & DataI & " and MAG80DAT.FTMOV00F.DTBOFM <= " & DataF & " " & _
                "and MAG80DAT.FTMOV00F.TDOCFM='B' and MAG80DAT.FTMOV00F.CDCFFM <> '" & Happy & "' " & _
                "GROUP BY MAG80DAT.FTMOV00F.CDCFFM,MAG80DAT.CGANA01J.CDPGCA,MAG80DAT.CGANA01J.CDAGCA,MAG80DAT.CGANA01J.CABICA,MAG80DAT.CGANA01J.CCABCA "

        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabCLI.Clear()
        myDA.Fill(tabCLI)
        odbcGalileo.Close()

        For Each rowCLI In tabCLI.Rows

            ClienteP = rowCLI("CLIENTE")
            AgenteP = rowCLI("COD_AGENTE")
            PagamentoP = rowCLI("CODICE_PAGAMENTO")
            AbiP = rowCLI("COD_ABI").ToString
            CabP = rowCLI("COD_CAB").ToString

            query = " select MAG80DAT.FTMOV00F.NRDFFM AS NUMDOCUMENTO, " &
                    " MAG80DAT.FTMOV00F.CDPAFM AS CODPAGAMENTO,MAG80DAT.FTMOV00F.CDAGFM AS CODAGENTE," &
                     " MAG80DAT.FTMOV00F.CDABFM AS CODABI,MAG80DAT.FTMOV00F.CDCAFM AS CODCAB" &
                    " from MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTMOV00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA" &
                    " where MAG80DAT.FTMOV00F.DTBOFM >= " & DataI & " And MAG80DAT.FTMOV00F.DTBOFM <= " & DataF &
                    " and MAG80DAT.FTMOV00F.TDOCFM='B' " &
                    " and MAG80DAT.CGANA01J.CONTCA='" & ClienteP & "'" &
                    " and (MAG80DAT.FTMOV00F.CDPAFM<>'" & PagamentoP & "'" &
                    " or MAG80DAT.FTMOV00F.CDAGFM<>'" & AgenteP & "' or MAG80DAT.FTMOV00F.CDABFM <> '" & AbiP & "' or MAG80DAT.FTMOV00F.CDCAFM <> '" & CabP & "')"

            query = query & " union all " & " select MAG80DAT.FTBKM00F.NRDFFM AS NUMDOCUMENTO, " &
                    " MAG80DAT.FTBKM00F.CDPAFM AS CODPAGAMENTO,MAG80DAT.FTBKM00F.CDAGFM AS CODAGENTE," &
                     " MAG80DAT.FTBKM00F.CDABFM AS CODABI,MAG80DAT.FTBKM00F.CDCAFM AS CODCAB" &
                    " from MAG80DAT.FTBKM00F LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTBKM00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA" &
                    " where MAG80DAT.FTBKM00F.DTBOFM >= " & DataI & " And MAG80DAT.FTBKM00F.DTBOFM <= " & DataF &
                    " and MAG80DAT.FTBKM00F.TDOCFM='B' " &
                    " and MAG80DAT.CGANA01J.CONTCA='" & ClienteP & "'" &
                    " and (MAG80DAT.FTBKM00F.CDPAFM<>'" & PagamentoP & "'" &
                    " or MAG80DAT.FTBKM00F.CDAGFM<>'" & AgenteP & "' or MAG80DAT.FTBKM00F.CDABFM <> '" & AbiP & "' or MAG80DAT.FTBKM00F.CDCAFM <> '" & CabP & "')"

            tmpTabControlli.Clear()
            odbcGalileo.Open()
            myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
            myDA.Fill(tmpTabControlli)
            odbcGalileo.Close()

            For Each row In tmpTabControlli.Rows

                Cliente = ClienteP & " " & Trim(NomeClienteP)

                Agente = ""

                'test
                Dim test As String = ""

                If Not IsDBNull(row("NUMDOCUMENTO")) Then
                    test = row("NUMDOCUMENTO")
                End If

                'fine test
                If Not IsDBNull(row("CODAGENTE")) Then
                    For Each rowAgente In tmpTabAgenti.Rows
                        If rowAgente("CODICE_AGENTE").ToString.Trim = row("CODAGENTE") Then
                            Agente = rowAgente("DESCRIZIONE_AGENTE").ToString.Trim
                            Exit For
                        End If
                    Next
                Else
                    Agente = ""
                End If

                Pagamento = ""

                If Not IsDBNull(row("CODPAGAMENTO")) Then
                    For Each rowPagamento In tmpTabPagamenti.Rows
                        If rowPagamento("CODICE_PAGAMENTO").ToString.Trim = row("CODPAGAMENTO") Then
                            Pagamento = rowPagamento("DESCRIZIONE_PAGAMENTO").ToString.Trim
                            Exit For
                        End If
                    Next
                Else
                    Pagamento = ""
                End If

                Abi = row("CODABI")
                Cab = row("CODCAB")

                DgvDati1.Rows.Add(New String() {Cliente, row("NUMDOCUMENTO"), Pagamento, Agente, Abi, Cab})
            Next
        Next

        Cursor = Cursors.Default

        If DgvDati1.Rows.Count = 0 Then
            MessageBox.Show("NESSUNA ANOMALIA RISCONTRATA")
        End If
    End Sub

    Public Sub DgvArchivio_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DgvArchivio.CurrentCellDirtyStateChanged
        If DgvArchivio.IsCurrentCellDirty Then
            DgvArchivio.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DgvArchivio_RowLeave(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvArchivio.RowLeave
        If DgvArchivio.Rows(e.RowIndex).Cells(0).Value <> Nothing And DgvArchivio.Rows(e.RowIndex).Cells(0).Value <> "CANCELLATO" Then

            Dim numdoc As String = ""
            Dim tmpData As String = ""

            If DgvArchivio.Rows(e.RowIndex).Cells(1).Value = Nothing Then
                flag_vettore = 0
            ElseIf DgvArchivio.Rows(e.RowIndex).Cells(1).Value = True Then
                flag_vettore = 1
            ElseIf DgvArchivio.Rows(e.RowIndex).Cells(1).Value = False Then
                flag_vettore = 0
            End If

            If DgvArchivio.Rows(e.RowIndex).Cells(2).Value = Nothing Then
                flag_cliente = 0
            ElseIf DgvArchivio.Rows(e.RowIndex).Cells(2).Value = True Then
                flag_cliente = 1
            ElseIf DgvArchivio.Rows(e.RowIndex).Cells(2).Value = False Then
                flag_cliente = 0
            End If

            If IsNumeric(DgvArchivio.Rows(e.RowIndex).Cells(0).Value) Then
                numdoc = DgvArchivio.Rows(e.RowIndex).Cells(0).Value
            Else
                MessageBox.Show("Numero documento non valido.", "Errore")
                If cnDB.State = ConnectionState.Open Then cnDB.Close()
                Exit Sub
            End If

            If DgvArchivio.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor <> Color.Orange Then

                If Not checkPresenzaDocumento(numdoc, e.RowIndex) Then

                    If flag_vettore = 0 And flag_cliente = 0 And Not DgvArchivio.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.Orange Then
                        MessageBox.Show("Selezionare se il documento è firmato dal cliente o meno.", "Errore")
                        If cnDB.State = ConnectionState.Open Then cnDB.Close()
                        Exit Sub
                    End If

                    Try
                        query = "SELECT MAG80DAT.FTMOV00F.DTBOFM as DtDDT " & _
                                "from MAG80DAT.FTMOV00F " & _
                                "where MAG80DAT.FTMOV00F.NRDFFM = " & "'" & numdoc & "' " & _
                                "and MAG80DAT.FTMOV00F.TDOCFM='B' " & _
                                "UNION ALL " & _
                                "SELECT MAG80DAT.FTBKM00F.DTBOFM as DtDDT " & _
                                "from MAG80DAT.FTBKM00F " & _
                                "where MAG80DAT.FTBKM00F.NRDFFM = " & "'" & numdoc & "' " & _
                                "and MAG80DAT.FTBKM00F.TDOCFM='B' "

                        odbcGalileo.ConnectionString = myConnectionStringGalileo
                        odbcGalileo.Open()
                        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
                        tabCLI.Clear()
                        myDA.Fill(tabCLI)
                        odbcGalileo.Close()
                    Catch
                        MessageBox.Show("Errore sull'interrogazione.", "Errore")
                        If cnDB.State = ConnectionState.Open Then cnDB.Close()
                        Exit Sub
                    End Try

                    For Each rowCLI In tabCLI.Rows
                        If Not IsDBNull(rowCLI("DtDDT")) Then
                            tmpData = rowCLI("DtDDT")
                            tmpData = tmpData.Substring(0, 4) & "/" & tmpData.Substring(4, 2) & "/" & tmpData.Substring(6, 2)
                            DgvArchivio.Rows(e.RowIndex).Cells(3).Value = tmpData
                        End If
                    Next

                    If DgvArchivio.Rows(e.RowIndex).Cells(3).Value = "" Then

                        DgvArchivio.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Red
                        DgvArchivio.Rows(e.RowIndex).Cells(1).Style.BackColor = Color.Red
                        DgvArchivio.Rows(e.RowIndex).Cells(2).Style.BackColor = Color.Red
                        DgvArchivio.Rows(e.RowIndex).Cells(3).Style.BackColor = Color.Red

                        MessageBox.Show("Documento non trovato!", "Errore!")
                    Else

1:                      cnDB.ConnectionString = myConnectionString

                        ' CONTROLLO CHE IL DOCUMENTO IN INSERIMENTO NON SIA GIA' PRESENTE IN ARCHIVIO
                        query = "SELECT * FROM DOCUMENTI WHERE NUMDOCUMENTO=" & numdoc

                        myDAccess.SelectCommand = New OleDbCommand(query, cnDB)

                        Dim tmpTabDocumenti As New DataTable

                        cnDB.Open()

                        If myDAccess.Fill(tmpTabDocumenti) = 0 Then
                            query = "INSERT INTO DOCUMENTI VALUES ('" & tmpData & "'," & numdoc & "," & _
                                    flag_vettore & "," & flag_cliente & ")"
                            myDAccess.InsertCommand = New OleDbCommand(query, cnDB)
                            myDAccess.InsertCommand.ExecuteNonQuery()
                            cnDB.Close()

                            DgvArchivio.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.LightGreen
                            DgvArchivio.Rows(e.RowIndex).Cells(0).ReadOnly = True

                            DgvArchivio.Rows(e.RowIndex).Cells(1).Style.BackColor = Color.LightGreen
                            DgvArchivio.Rows(e.RowIndex).Cells(2).Style.BackColor = Color.LightGreen
                            DgvArchivio.Rows(e.RowIndex).Cells(3).Style.BackColor = Color.LightGreen
                            maxRiga += 1
                            lastNrDoc = numdoc

                        ElseIf flag_cliente = 0 And flag_vettore = 0 Then

                            Dim result As Integer = MessageBox.Show("Entrambi i flag sono deselezionati, procedo alla cancellazione?", "Attenzione", MessageBoxButtons.OKCancel)
                            If result = DialogResult.Cancel Then

                                DgvArchivio.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.Yellow
                                DgvArchivio.Rows(e.RowIndex).Cells(1).Style.BackColor = Color.Yellow
                                DgvArchivio.Rows(e.RowIndex).Cells(2).Style.BackColor = Color.Yellow
                                DgvArchivio.Rows(e.RowIndex).Cells(3).Style.BackColor = Color.Yellow

                                MessageBox.Show("Operazione annullata.")
                            ElseIf result = DialogResult.OK Then
                                Try
                                    query = "DELETE FROM DOCUMENTI WHERE NUMDOCUMENTO=" & numdoc
                                    myDAccess.DeleteCommand = New OleDbCommand(query, cnDB)
                                    myDAccess.DeleteCommand.ExecuteNonQuery()
                                    cnDB.Close()

                                    For ii As Integer = 0 To DgvArchivio.Rows.Count - 1
                                        If DgvArchivio.Rows(ii).Cells(0).Value = numdoc Then
                                            DgvArchivio.Rows(ii).Cells(0).Value = "CANCELLATO"
                                            DgvArchivio.Rows(ii).Cells(1).Value = 0
                                            DgvArchivio.Rows(ii).Cells(2).Value = 0
                                            DgvArchivio.Rows(ii).Cells(3).Value = "CANCELLATO"

                                            DgvArchivio.Rows(ii).Cells(0).Style.BackColor = Color.Gray
                                            DgvArchivio.Rows(ii).Cells(1).Style.BackColor = Color.Gray
                                            DgvArchivio.Rows(ii).Cells(2).Style.BackColor = Color.Gray
                                            DgvArchivio.Rows(ii).Cells(3).Style.BackColor = Color.Gray
                                        End If
                                    Next
                                Catch
                                End Try
                                MessageBox.Show("Cancellazione completata")
                            End If

                        Else
                            query = "UPDATE DOCUMENTI SET FIRMATAVETTORE=" & flag_vettore & ", FIRMATACLIENTE =" & flag_cliente & "" & _
                                    " WHERE NUMDOCUMENTO=" & numdoc
                            myDAccess.UpdateCommand = New OleDbCommand(query, cnDB)
                            myDAccess.UpdateCommand.ExecuteNonQuery()
                            cnDB.Close()

                            DgvArchivio.Rows(e.RowIndex).Cells(0).Style.BackColor = Color.LightGreen
                            DgvArchivio.Rows(e.RowIndex).Cells(1).Style.BackColor = Color.LightGreen
                            DgvArchivio.Rows(e.RowIndex).Cells(2).Style.BackColor = Color.LightGreen
                            DgvArchivio.Rows(e.RowIndex).Cells(3).Style.BackColor = Color.LightGreen

                            lastNrDoc = numdoc

                            MessageBox.Show("Documento già presente aggiornato.")
                        End If

                        flag_cliente = "0"
                        flag_vettore = "0"
                    End If
                End If
            Else
                GoTo 1
            End If
        End If
    End Sub

    Private Sub btnControllaPresenza_Click(sender As System.Object, e As System.EventArgs) Handles btnControllaPresenza.Click

        Dim Cliente As String = ""
        Dim Destinatario As String = ""
        Dim mancanti001 As Integer = 0
        Dim mancantiAltri As Integer = 0
        Dim Is1 As Integer = 0
        Dim Is2 As Integer = 0
        Dim Is3 As Integer = 0
        Dim Ibo As Integer = 0

        Dim docDA As String = ""
        Dim docA As String = ""
        Dim whereDocAct As String = ""
        Dim whereDocStorico As String = ""

        If checkNumDoc.Checked Then
            If IsNumeric(txtDocFiltroDA.Text) AndAlso IsNumeric(txtDocFiltroA.Text) Then
                docDA = txtDocFiltroDA.Text
                docA = txtDocFiltroA.Text

                whereDocAct = " AND MAG80DAT.FTMOV00F.NRDFFM >= " & docDA & " AND MAG80DAT.FTMOV00F.NRDFFM <= " & docA & " "

                whereDocStorico = " AND MAG80DAT.FTBKM00F.NRDFFM >= " & docDA & " AND MAG80DAT.FTBKM00F.NRDFFM <= " & docA & " "

            Else
                MessageBox.Show("Dati errati in filtro numero documento!", "Attenzione")
                Exit Sub
            End If
        End If

        Cursor = Cursors.WaitCursor

        dgvCheckPresenzaDoc.Visible = True
        dgvCheckPresenzaDoc1.Visible = False
        CmbSTATO.Visible = False
        CmbBOLLETTARIO.Visible = False

        dgvCheckPresenzaDoc.Rows.Clear()
        CmbSTATO.Items.Clear()
        CmbBOLLETTARIO.Items.Clear()

        DataI = DatDA.Value.Year.ToString
        If DatDA.Value.Month.ToString.Length > 1 Then
            DataI = DataI & DatDA.Value.Month.ToString
        Else
            DataI = DataI & "0" & DatDA.Value.Month.ToString
        End If
        If DatDA.Value.Day.ToString.Length > 1 Then
            DataI = DataI & DatDA.Value.Day.ToString
        Else
            DataI = DataI & "0" & DatDA.Value.Day.ToString
        End If

        DataF = DatA.Value.Year.ToString
        If DatA.Value.Month.ToString.Length > 1 Then
            DataF = DataF & DatA.Value.Month.ToString
        Else
            DataF = DataF & "0" & DatA.Value.Month.ToString
        End If
        If DatA.Value.Day.ToString.Length > 1 Then
            DataF = DataF & DatA.Value.Day.ToString
        Else
            DataF = DataF & "0" & DatA.Value.Day.ToString
        End If

        query = "SELECT * FROM DOCUMENTI"

        cnDB.ConnectionString = myConnectionString
        myDAccess.SelectCommand = New OleDbCommand(query, cnDB)

        tabDocumentiArchivio.Clear()

        cnDB.Open()
        myDAccess.Fill(tabDocumentiArchivio)
        cnDB.Close()

        getBollettario(DataI, DataF)

        For Each rowBOL In tabBOL.Rows
            query = "select MAG80DAT.FTMOV00F.NRDFFM AS NrDoc, " & _
                    "MAG80DAT.FTMOV00F.CDCFFM AS CLIENTE, " & _
                    "MAG80DAT.CGANA01J.DSCOCP AS RAGCLI, " & _
                    "MAG80DAT.FTMOV00F.CSPEFM AS DESTINATARIO, " & _
                    "MAG80DAT.FTMOV00F.DTBOFM AS DATA_BOLLA, " & _
                    "MAG80DAT.FTMOV00F.NRBOFM AS NR_BOLLA, " & _
                    "DEST.DSCOCP AS RAGDES, " & _
                    "trim(MAG80DAT.FTMOV00F.CDSPFM) AS PrimoVet , " & _
                    "trim(MAG80DAT.FTMOV00F.CDVEFM) AS SecondoVet " & _
                    "from (MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.CGANA01J ON " & _
                    "MAG80DAT.FTMOV00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA) " & _
                    "LEFT JOIN MAG80DAT.CGANA01J DEST ON MAG80DAT.FTMOV00F.CSPEFM=DEST.CONTCA " & _
                    "where MAG80DAT.FTMOV00F.CDBOFM = " & "'" & rowBOL("BOLLETTARIO") & "' " & _
                    "and MAG80DAT.FTMOV00F.DTBOFM >= " & "'" & DataI & "' " & _
                    "and MAG80DAT.FTMOV00F.DTBOFM <= " & "'" & DataF & "' " & _
                    "and MAG80DAT.FTMOV00F.TDOCFM='B' " & whereDocAct & _
                    "UNION ALL " & _
                    "select MAG80DAT.FTBKM00F.NRDFFM AS NrDoc, " & _
                    "MAG80DAT.FTBKM00F.CDCFFM AS CLIENTE, " & _
                    "MAG80DAT.CGANA01J.DSCOCP AS RAGCLI, " & _
                    "MAG80DAT.FTBKM00F.CSPEFM AS DESTINATARIO, " & _
                    "MAG80DAT.FTBKM00F.DTBOFM AS DATA_BOLLA, " & _
                    "MAG80DAT.FTBKM00F.NRBOFM AS NR_BOLLA, " & _
                    "DEST.DSCOCP AS RAGDES, " & _
                    "trim(MAG80DAT.FTBKM00F.CDSPFM) AS PrimoVet, " & _
                    "trim(MAG80DAT.FTBKM00F.CDVEFM) AS SecondoVet " & _
                    "from (MAG80DAT.FTBKM00F LEFT JOIN MAG80DAT.CGANA01J ON " & _
                    "MAG80DAT.FTBKM00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA) " & _
                    "LEFT JOIN MAG80DAT.CGANA01J DEST ON MAG80DAT.FTBKM00F.CSPEFM=DEST.CONTCA " & _
                    "where MAG80DAT.FTBKM00F.CDBOFM = " & "'" & rowBOL("BOLLETTARIO") & "' " & _
                    "and MAG80DAT.FTBKM00F.DTBOFM >= " & "'" & DataI & "' " & _
                    "and MAG80DAT.FTBKM00F.DTBOFM <= " & "'" & DataF & "' " & _
                    "and MAG80DAT.FTBKM00F.TDOCFM='B' " & whereDocStorico & _
                    "ORDER BY NrDoc"

            myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
            tabDDT.Clear()
            odbcGalileo.Open()
            myDA.Fill(tabDDT)
            odbcGalileo.Close()

            For Each rowDDT In tabDDT.Rows
                RagsPrimoVet = ""
                RagsSecondoVet = ""

                If Not IsDBNull(rowDDT("PrimoVet")) And Not IsDBNull(rowDDT("SecondoVet")) Then
                    findVet(rowDDT("PrimoVet"), rowDDT("SecondoVet"))
                ElseIf Not IsDBNull(rowDDT("PrimoVet")) And IsDBNull(rowDDT("SecondoVet")) Then
                    findVet(rowDDT("PrimoVet"), "")
                ElseIf IsDBNull(rowDDT("PrimoVet")) And IsDBNull(rowDDT("SecondoVet")) Then
                    findVet("", rowDDT("SecondoVet"))
                End If
                Cliente = rowDDT("CLIENTE") & " - " & rowDDT("RAGCLI")
                Destinatario = rowDDT("DESTINATARIO") & " - " & rowDDT("RAGDES")
                dgvCheckPresenzaDoc.Rows.Add(New String() {rowDDT("NrDoc"), rowBOL("BOLLETTARIO"), "", Cliente, Destinatario, RagsPrimoVet, RagsSecondoVet, rowDDT("NR_BOLLA"), rowDDT("DATA_BOLLA")})

                Ibo = 0
                Do While Ibo < CmbBOLLETTARIO.Items.Count AndAlso CmbBOLLETTARIO.Items(Ibo).ToString <> rowBOL("BOLLETTARIO")
                    Ibo = Ibo + 1
                Loop
                If Ibo = CmbBOLLETTARIO.Items.Count Then
                    CmbBOLLETTARIO.Items.Add(rowBOL("BOLLETTARIO"))
                End If
            Next
        Next

        If dgvCheckPresenzaDoc.Rows.Count >= 1 Then

            btnCreaReport.Visible = True

            For ii As Integer = 0 To dgvCheckPresenzaDoc.Rows.Count - 1

                Dim tmpNumDoc As String = ""

                If dgvCheckPresenzaDoc.Rows(ii).Cells(0).Value <> "" Then
                    tmpNumDoc = dgvCheckPresenzaDoc.Rows(ii).Cells(0).Value
                End If

                For Each rowDocumentiArchivio In tabDocumentiArchivio.Rows
                    If tmpNumDoc <> "" Then
                        If rowDocumentiArchivio("NumDocumento") = tmpNumDoc Then
                            If rowDocumentiArchivio("FirmataVettore") Then
                                dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(1).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(2).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(3).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(4).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(5).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(6).Style.BackColor = Color.Orange
                                dgvCheckPresenzaDoc.Rows(ii).Cells(2).Value = "SOLO INTERNA"

                                Is1 = Is1 + 1
                                If Is1 = 1 Then
                                    CmbSTATO.Items.Add(dgvCheckPresenzaDoc.Rows(ii).Cells(2).Value)
                                End If
                            End If
                            If rowDocumentiArchivio("FirmataCliente") Then
                                dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor = Color.LightGreen
                                dgvCheckPresenzaDoc.Rows(ii).Cells(1).Style.BackColor = Color.LightGreen
                                dgvCheckPresenzaDoc.Rows(ii).Cells(2).Style.BackColor = Color.LightGreen
                                dgvCheckPresenzaDoc.Rows(ii).Cells(3).Style.BackColor = Color.LightGreen
                                dgvCheckPresenzaDoc.Rows(ii).Cells(4).Style.BackColor = Color.LightGreen
                                dgvCheckPresenzaDoc.Rows(ii).Cells(5).Value = ""
                                dgvCheckPresenzaDoc.Rows(ii).Cells(5).Style.BackColor = Color.LightGreen
                                dgvCheckPresenzaDoc.Rows(ii).Cells(6).Value = ""
                                dgvCheckPresenzaDoc.Rows(ii).Cells(6).Style.BackColor = Color.LightGreen

                                dgvCheckPresenzaDoc.Rows(ii).Cells(2).Value = "CLIENTE PRESENTE"

                                Is2 = Is2 + 1
                                If Is2 = 1 Then
                                    CmbSTATO.Items.Add(dgvCheckPresenzaDoc.Rows(ii).Cells(2).Value)
                                End If
                            End If
                        End If
                    End If
                Next

                If dgvCheckPresenzaDoc.Rows(ii).Cells(0).Value <> Nothing And dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor <> Color.LightGreen And dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor <> Color.Orange Then
                    dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(1).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(2).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(3).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(4).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(5).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(6).Style.BackColor = Color.Red
                    dgvCheckPresenzaDoc.Rows(ii).Cells(2).Value = "MANCANTE"

                    Is3 = Is3 + 1
                    If Is3 = 1 Then
                        CmbSTATO.Items.Add(dgvCheckPresenzaDoc.Rows(ii).Cells(2).Value)
                    End If

                    If dgvCheckPresenzaDoc.Rows(ii).Cells(1).Value <> Nothing AndAlso dgvCheckPresenzaDoc.Rows(ii).Cells(1).Value.ToString = "001" Then
                        mancanti001 += 1
                    ElseIf dgvCheckPresenzaDoc.Rows(ii).Cells(1).Value <> Nothing Then
                        mancantiAltri += 1
                    End If
                    contaDocMancanti += 1
                ElseIf dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor = Color.LightGreen Or dgvCheckPresenzaDoc.Rows(ii).Cells(0).Style.BackColor = Color.Orange Then
                    contaDocPresenti += 1
                End If
            Next

            If CmbBOLLETTARIO.Items.Count > 1 Then
                CmbBOLLETTARIO.Visible = True
                CmbBOLLETTARIO.Items.Add("Tutti")
            End If
            If CmbSTATO.Items.Count > 1 Then
                CmbSTATO.Visible = True
                CmbSTATO.Items.Add("Tutti")
            End If
        Else
            btnCreaReport.Visible = False
        End If

        lblMancanti001.Text = mancanti001.ToString
        lblMancantiAltri.Text = mancantiAltri.ToString

        lblDescrMancanti001.Visible = True
        lblMancanti001.Visible = True

        lblDescrMancantiAltri.Visible = True
        lblMancantiAltri.Visible = True

        Cursor = Cursors.Default

        MessageBox.Show("Operazione completata. Numero documenti mancanti: " & contaDocMancanti & ". Numero documenti presenti:" & contaDocPresenti & ".")
        contaDocMancanti = 0
        contaDocPresenti = 0
        mancanti001 = 0
        mancantiAltri = 0
    End Sub

    Private Sub getBollettario(ByVal _dataInizio As String, ByVal _dataFine As String)
        Dim ii As Integer

        query = "select distinct MAG80DAT.FTMOV00F.CDBOFM AS BOLLETTARIO " &
                "from MAG80DAT.FTMOV00F " &
                "where MAG80DAT.FTMOV00F.DTBOFM >= " & _dataInizio & " " &
                "and MAG80DAT.FTMOV00F.DTBOFM <= " & _dataFine & " " &
                "and MAG80DAT.FTMOV00F.TDOCFM='B' " &
                "UNION ALL " &
                "select distinct MAG80DAT.FTBKM00F.CDBOFM AS BOLLETTARIO " &
                "from MAG80DAT.FTBKM00F " &
                "where MAG80DAT.FTBKM00F.DTBOFM >= " & _dataInizio & " " &
                "and MAG80DAT.FTBKM00F.DTBOFM <= " & _dataFine & " " &
                "and MAG80DAT.FTBKM00F.TDOCFM='B' " &
                "order BY 1"

        odbcGalileo.ConnectionString = myConnectionStringGalileo
        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabDDT.Clear()
        tabBOL.Clear()
        myDA.Fill(tabDDT)
        odbcGalileo.Close()

        tabBOL = tabDDT.Clone
        Bollettario = " "
        ii = 0
        For Each rowDDT In tabDDT.Rows
            If rowDDT("BOLLETTARIO") <> Bollettario Then
                tabBOL.ImportRow(tabDDT.Rows(ii))
                Bollettario = rowDDT("BOLLETTARIO")
            End If
            ii = ii + 1
        Next
        tabDDT.Clear()
        Bollettario = " "
    End Sub

    Private Function checkPresenzaDocumento(_numdoc As String, _rowIndex As Integer) As Boolean
        ' SE IL NUMERO DOCUMENTO E' GIA' PRESENTE, MOSTRO I DATI PRESENTI PRIMA DELLA MODIFICA
        Dim docPresente As Boolean = False

        query = "SELECT * FROM DOCUMENTI WHERE NUMDOCUMENTO=" & _numdoc

        Dim tmpTabDocumenti As New DataTable

        cnDB.ConnectionString = myConnectionString

        myDAccess.SelectCommand = New OleDbCommand(query, cnDB)

        cnDB.Open()

        If myDAccess.Fill(tmpTabDocumenti) > 0 Then
            docPresente = True
            For Each tmpRow As DataRow In tmpTabDocumenti.Rows
                If tmpRow("FirmataVettore") Then DgvArchivio.Rows(_rowIndex).Cells(1).Value = True Else DgvArchivio.Rows(_rowIndex).Cells(1).Value = False
                If tmpRow("FirmataCliente") Then DgvArchivio.Rows(_rowIndex).Cells(2).Value = True Else DgvArchivio.Rows(_rowIndex).Cells(2).Value = False
                DgvArchivio.Rows(_rowIndex).Cells(3).Value = tmpRow("Data").ToString.Replace(" 00:00:00", "")
            Next
            DgvArchivio.Rows(_rowIndex).Cells(0).Style.BackColor = Color.Orange
            DgvArchivio.Rows(_rowIndex).Cells(1).Style.BackColor = Color.Orange
            DgvArchivio.Rows(_rowIndex).Cells(2).Style.BackColor = Color.Orange
            DgvArchivio.Rows(_rowIndex).Cells(3).Style.BackColor = Color.Orange
        End If
        cnDB.Close()

        Return docPresente
    End Function

    Private Sub btnCheckAnagrafiche_Click(sender As System.Object, e As System.EventArgs) Handles btnCheckAnagrafiche.Click
        Dim SpeseBolli As String = " "
        Dim SpeseIncasso As String = " "
        Dim AddebitoIvaOmaggi As String = " "
        Dim TipoFatturazione As String = " "
        Dim OrdinamentoFatture As String = " "
        Dim FiltroFatturazione As String = " "
        Dim RaggruppamentoFatture As String = " "
        Dim CodRaggruppamentoFatture As String = " "
        Dim FatturazioneGruppoArticolo As String = " "
        Dim Anomalia As Boolean
        Dim GriSpeseBolli As String
        Dim GriDestinatario As String
        Dim GriSpeseIncasso As String
        Dim GriAddebitoIvaOmaggi As String
        Dim GriTipoFatturazione As String
        Dim GriOrdinamentoFatture As String
        Dim GriFiltroFatturazione As String
        Dim GriRaggruppamentoFatture As String
        Dim GriCodRaggruppamentoFatture As String
        Dim GriFatturazioneGruppoArticolo As String
        Dim GriClienteEmissFatture As String

        Cursor = Cursors.WaitCursor

        query = "SELECT trim(MAG80DAT.CGANA00F.SPBLCA) AS SPESE_BOLLI, " & _
                "trim(MAG80DAT.CGANA02F.SPINAC) AS SPESE_INCASSO, " & _
                "trim(MAG80DAT.CGANA02F.FL09AC) AS ADDEBITO_IVA_OMAGGI, " & _
                "trim(MAG80DAT.CGANA03F.TPFTAD) AS TIPO_FATTURAZIONE, " & _
                "trim(MAG80DAT.CGANA03F.ORFTAD) AS ORDINAMENTO_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.TPDAAD) AS FILTRO_FATTURAZIONE, " & _
                "trim(MAG80DAT.CGANA03F.FFRGAD) AS RAGGRUPPAMENTO_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.RIFTAD) AS COD_RAGGRUPPAMENTO_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.FLFTAD) AS FATTURAZIONE_GRUPPO_ARTICOLO " & _
                "FROM ((MAG80DAT.CGANA00F LEFT JOIN MAG80DAT.CGANA02F ON MAG80DAT.CGANA00F.CONTCA =MAG80DAT.CGANA02F.CONTAC) " & _
                "LEFT JOIN MAG80DAT.CGANA03F ON MAG80DAT.CGANA00F.CONTCA = MAG80DAT.CGANA03F.CONTAD) " & _
                "LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.CGANA00F.CONTCA = MAG80DAT.CGANA01J.CONTCA " & _
                "WHERE MAG80DAT.CGANA00F.CONTCA = '" & Happy & "'"

        odbcGalileo.ConnectionString = myConnectionStringGalileo
        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabAnag1503.Clear()
        myDA.Fill(tabAnag1503)

        For Each rowAnag1503 In tabAnag1503.Rows
            SpeseBolli = rowAnag1503("SPESE_BOLLI")
            SpeseIncasso = rowAnag1503("SPESE_INCASSO")
            AddebitoIvaOmaggi = rowAnag1503("ADDEBITO_IVA_OMAGGI")
            TipoFatturazione = rowAnag1503("TIPO_FATTURAZIONE")
            OrdinamentoFatture = rowAnag1503("ORDINAMENTO_FATTURE")
            FiltroFatturazione = rowAnag1503("FILTRO_FATTURAZIONE")
            RaggruppamentoFatture = rowAnag1503("RAGGRUPPAMENTO_FATTURE")
            CodRaggruppamentoFatture = rowAnag1503("COD_RAGGRUPPAMENTO_FATTURE")
            FatturazioneGruppoArticolo = rowAnag1503("FATTURAZIONE_GRUPPO_ARTICOLO")
        Next

        query = "SELECT  trim(MAG80DAT.CGANA00F.CONTCA) AS COD_DEST, " & _
                "trim(MAG80DAT.CGANA01J.DSCOCP) AS DESTINATARIO, " & _
                "trim(MAG80DAT.CGANA02F.SPINAC) AS SPESE_INCASSO, " & _
                "trim(MAG80DAT.CGANA02F.FL09AC) AS ADDEBITO_IVA_OMAGGI, " & _
                "trim(MAG80DAT.CGANA03F.TPFTAD) AS TIPO_FATTURAZIONE, " & _
                "trim(MAG80DAT.CGANA00F.SPBLCA) AS SPESE_BOLLI, " & _
                "trim(MAG80DAT.CGANA03F.CDGCAD) AS CLIENTE_EMISS_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.ORFTAD) AS ORDINAMENTO_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.TPDAAD) AS FILTRO_FATTURAZIONE, " & _
                "trim(MAG80DAT.CGANA03F.FFRGAD) AS RAGGRUPPAMENTO_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.RIFTAD) AS COD_RAGGRUPPAMENTO_FATTURE, " & _
                "trim(MAG80DAT.CGANA03F.FLFTAD) AS FATTURAZIONE_GRUPPO_ARTICOLO " & _
                "FROM ((MAG80DAT.CGANA00F LEFT JOIN MAG80DAT.CGANA02F ON MAG80DAT.CGANA00F.CONTCA =MAG80DAT.CGANA02F.CONTAC) " & _
                "LEFT JOIN MAG80DAT.CGANA03F ON MAG80DAT.CGANA00F.CONTCA = MAG80DAT.CGANA03F.CONTAD) " & _
                "LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.CGANA00F.CONTCA = MAG80DAT.CGANA01J.CONTCA " & _
                "WHERE MAG80DAT.CGANA00F.CONTCA like '1503%'  and (trim(MAG80DAT.CGANA00F.SPBLCA)  <>'' or " & _
                "trim(MAG80DAT.CGANA02F.SPINAC) <> ''  or trim(MAG80DAT.CGANA02F.FL09AC) <> '' or " & _
                "trim(MAG80DAT.CGANA03F.CDGCAD) <>'" & Happy & "' or trim(MAG80DAT.CGANA03F.ORFTAD) <>'' or " & _
                "trim(MAG80DAT.CGANA03F.TPFTAD) <> ''  or 	trim(MAG80DAT.CGANA03F.FFRGAD) <>'' or " & _
                "trim(MAG80DAT.CGANA03F.RIFTAD) <>' ' or trim(MAG80DAT.CGANA03F.FLFTAD) <>'') " & _
                "ORDER BY 1 "

        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabAnag1503.Clear()
        Dim i As Integer = myDA.Fill(tabAnag1503)
        odbcGalileo.Close()

        If i > 0 Then

            dgvAnagrafiche1503.Rows.Clear()

            For Each rowAnag1503 In tabAnag1503.Rows
                Anomalia = False
                GriSpeseBolli = " "
                GriDestinatario = " "
                GriSpeseIncasso = " "
                GriAddebitoIvaOmaggi = " "
                GriTipoFatturazione = " "
                GriOrdinamentoFatture = " "
                GriFiltroFatturazione = " "
                GriClienteEmissFatture = " "
                GriRaggruppamentoFatture = " "
                GriCodRaggruppamentoFatture = " "
                GriFatturazioneGruppoArticolo = " "

                If rowAnag1503("CLIENTE_EMISS_FATTURE") <> Happy Then
                    GriClienteEmissFatture = rowAnag1503("CLIENTE_EMISS_FATTURE")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("SPESE_BOLLI")) And rowAnag1503("SPESE_BOLLI") <> SpeseBolli Then
                    GriSpeseIncasso = rowAnag1503("SPESE_BOLLI")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("SPESE_INCASSO")) And rowAnag1503("SPESE_INCASSO") <> SpeseIncasso Then
                    GriSpeseIncasso = rowAnag1503("SPESE_INCASSO")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("ADDEBITO_IVA_OMAGGI")) And rowAnag1503("ADDEBITO_IVA_OMAGGI") <> AddebitoIvaOmaggi Then
                    GriAddebitoIvaOmaggi = rowAnag1503("ADDEBITO_IVA_OMAGGI")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("TIPO_FATTURAZIONE")) And rowAnag1503("TIPO_FATTURAZIONE") <> TipoFatturazione Then
                    GriTipoFatturazione = rowAnag1503("TIPO_FATTURAZIONE")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("ORDINAMENTO_FATTURE")) And rowAnag1503("ORDINAMENTO_FATTURE") <> OrdinamentoFatture Then
                    GriOrdinamentoFatture = rowAnag1503("ORDINAMENTO_FATTURE")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("FILTRO_FATTURAZIONE")) And rowAnag1503("FILTRO_FATTURAZIONE") <> FiltroFatturazione Then
                    GriFiltroFatturazione = rowAnag1503("FILTRO_FATTURAZIONE")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("RAGGRUPPAMENTO_FATTURE")) And rowAnag1503("RAGGRUPPAMENTO_FATTURE") <> RaggruppamentoFatture Then
                    GriRaggruppamentoFatture = rowAnag1503("RAGGRUPPAMENTO_FATTURE")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("COD_RAGGRUPPAMENTO_FATTURE")) And rowAnag1503("COD_RAGGRUPPAMENTO_FATTURE") <> CodRaggruppamentoFatture Then
                    GriCodRaggruppamentoFatture = rowAnag1503("COD_RAGGRUPPAMENTO_FATTURE")
                    Anomalia = True
                End If
                If Not IsDBNull(rowAnag1503("FATTURAZIONE_GRUPPO_ARTICOLO")) And rowAnag1503("FATTURAZIONE_GRUPPO_ARTICOLO") <> FatturazioneGruppoArticolo Then
                    GriFatturazioneGruppoArticolo = rowAnag1503("FATTURAZIONE_GRUPPO_ARTICOLO")
                    Anomalia = True
                End If

                If Anomalia = True Then
                    GriDestinatario = rowAnag1503("COD_DEST") & " - " & rowAnag1503("DESTINATARIO")
                    dgvAnagrafiche1503.Rows.Add(New String() {GriDestinatario, GriClienteEmissFatture, GriTipoFatturazione, GriFatturazioneGruppoArticolo, GriRaggruppamentoFatture, GriCodRaggruppamentoFatture, GriFiltroFatturazione, GriOrdinamentoFatture, GriSpeseBolli, GriSpeseIncasso})
                End If
                Cursor = Cursors.Default
            Next
        Else
            MessageBox.Show("NESSUNA ANOMALIA RISCONTRATA")
            Cursor = Cursors.Default
        End If
    End Sub

    Private Sub findVet(_primoVet As String, _secondoVet As String)
        RagsPrimoVet = ""
        RagsSecondoVet = ""

        For Each rowVettori In tabVettori.Rows
            If rowVettori("CODICE_VETTORE") = _primoVet Then
                RagsPrimoVet = rowVettori("Vettore")
            ElseIf rowVettori("CODICE_VETTORE") = _secondoVet Then
                RagsSecondoVet = rowVettori("Vettore")
            End If
        Next
    End Sub

    Private Sub btnCheckMov_Click(sender As System.Object, e As System.EventArgs) Handles btnCheckMov.Click
        Cursor = Cursors.WaitCursor

        query = " SELECT MAG80DAT.FTMOV00F.NRDFFM AS NUMERO_DOC, " &
                " MAG80DAT.FTMOV00F.DTBOFM AS DATA_BOLLA, " &
                " trim(MAG80DAT.CGANA01J.DSCOCP) AS CLIENTE, " &
                " MAG80DAT.FTMOV01F.TIMOFM AS TIPO_MOV, " &
                " MAG80DAT.FTMOV01F.ALIVFM AS COD_IVA " &
                " FROM (MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.FTMOV01F ON (MAG80DAT.FTMOV00F.CDDTFM = MAG80DAT.FTMOV01F.CDDTFM) " &
                " AND (MAG80DAT.FTMOV00F.TDOCFM = MAG80DAT.FTMOV01F.TDOCFM) AND (MAG80DAT.FTMOV00F.NRDFFM = MAG80DAT.FTMOV01F.NRDFFM)) " &
                " LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTMOV00F.CDCFFM = MAG80DAT.CGANA01J.CONTCA " &
                " WHERE (TIMOFM='29' AND ALIVFM<>'E6')  OR  (TIMOFM<>'29' AND ALIVFM='E6') "



        odbcGalileo.ConnectionString = myConnectionStringGalileo
        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabCheckMovimenti.Clear()
        Dim i As Integer = myDA.Fill(tabCheckMovimenti)
        odbcGalileo.Close()

        dgvCheckMovimenti.Rows.Clear()

        If i > 0 Then
            Dim tmpNumDoc As String = ""
            Dim tmpDataBolla As String = ""
            Dim tmpCliente As String = ""
            Dim tmpTipoMov As String = ""
            Dim tmpCodIva As String = ""

            For Each rowCheckMovimenti In tabCheckMovimenti.Rows

                If Not IsDBNull(rowCheckMovimenti("NUMERO_DOC")) Then
                    tmpNumDoc = rowCheckMovimenti("NUMERO_DOC")
                End If

                If Not IsDBNull(rowCheckMovimenti("DATA_BOLLA")) Then
                    tmpDataBolla = rowCheckMovimenti("DATA_BOLLA")
                End If

                If Not IsDBNull(rowCheckMovimenti("CLIENTE")) Then
                    tmpCliente = rowCheckMovimenti("CLIENTE")
                End If

                If Not IsDBNull(rowCheckMovimenti("TIPO_MOV")) Then
                    tmpTipoMov = rowCheckMovimenti("TIPO_MOV")
                End If

                If Not IsDBNull(rowCheckMovimenti("COD_IVA")) Then
                    tmpCodIva = rowCheckMovimenti("COD_IVA")
                End If

                dgvCheckMovimenti.Rows.Add(New String() {tmpNumDoc, tmpDataBolla, tmpCliente, tmpTipoMov, tmpCodIva})

            Next
            Cursor = Cursors.Default
        Else
            Cursor = Cursors.Default
            MessageBox.Show("NESSUNA ANOMALIA RISCONTRATA")
        End If

    End Sub

    Private Sub btnInserisci_Click(sender As System.Object, e As System.EventArgs) Handles btnInserisci.Click

        DgvArchivio.Rows.Clear()

        Dim docDa As String = ""
        Dim docA As String = ""
        Dim interno As Boolean
        Dim esterno As Boolean

        If IsNumeric(txtDaDoc.Text) AndAlso IsNumeric(txtADoc.Text) Then
            docDa = txtDaDoc.Text
            docA = txtADoc.Text

            If Convert.ToInt32(docDa) <= Convert.ToInt32(docA) Then
                If (Convert.ToInt32(docA) - Convert.ToInt32(docDa)) < 500 Then
                    If checkInterno.Checked Or checkEsterno.Checked Then
                        If checkInterno.Checked Then interno = True Else interno = False
                        If checkEsterno.Checked Then esterno = True Else esterno = False

                        Dim maxiDRow As Integer = DgvArchivio.Rows.Count - 1

                        If srcViaNumDoc Then
                            ' FOR DA NUMDOC MINIMO A NUMDOC MASSIMO, POI LEGGO I NUMDOC DALLA GRIGLIA E LI PROCESSO IN AS
                            For ii As Integer = Convert.ToInt32(docDa) To Convert.ToInt32(docA)
                                DgvArchivio.Rows.Add(New String() {ii.ToString, interno, esterno, "-", "X"})
                                insertIntoArchivio(maxiDRow, esterno, interno)
                            Next
                        Else
                            ' PER OGNI NUMBOLLA ESTRAGGO IL SUO NUMDOC E LO INSERISCO IN GRIGLIA, AL TERMINE LI PROCESSO TUTTI IN AS
                            findNumDoc(docDa, docA)

                            For Each TMProwBol In TMPtabBol.Rows
                                DgvArchivio.Rows.Add(New String() {TMProwBol("NrDOC"), interno, esterno, "-", "X"})
                                insertIntoArchivio(maxiDRow, esterno, interno)
                            Next
                        End If
                        MessageBox.Show("Inserimento di massa completato.")
                    Else
                        MessageBox.Show("Attenzione: Specificare se il documento è interno o esterno.", "Errore")
                    End If
                Else
                    MessageBox.Show("Attenzione: stai cercando di inserire troppi documenti!.", "Errore")
                End If
            Else
                MessageBox.Show("Attenzione: il numero documento DA è maggiore del numero documento A.", "Errore")
            End If
        Else
            MessageBox.Show("Attenzione: i valori inseriti non sono corretti.", "Errore")
        End If
    End Sub

    Private Function insertIntoArchivio(maxiDRow As Integer, _flagCliente As Boolean, _flagVettore As Boolean) As Integer

        For ii As Integer = maxiDRow To DgvArchivio.Rows.Count - 1
            Dim numdoc As String = DgvArchivio.Rows(ii).Cells(0).Value
            Dim tmpdata As String = ""
            If numdoc <> Nothing Then

                Try
                    query = "SELECT MAG80DAT.FTMOV00F.DTBOFM as DtDDT " & _
                            "from MAG80DAT.FTMOV00F " & _
                            "where MAG80DAT.FTMOV00F.NRDFFM = " & "'" & numdoc & "' " & _
                            "and MAG80DAT.FTMOV00F.TDOCFM='B' " & _
                            "UNION ALL " & _
                            "SELECT MAG80DAT.FTBKM00F.DTBOFM as DtDDT " & _
                            "from MAG80DAT.FTBKM00F " & _
                            "where MAG80DAT.FTBKM00F.NRDFFM = " & "'" & numdoc & "' " & _
                            "and MAG80DAT.FTBKM00F.TDOCFM='B' "

                    odbcGalileo.ConnectionString = myConnectionStringGalileo
                    odbcGalileo.Open()
                    myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
                    tabCLI.Clear()
                    myDA.Fill(tabCLI)
                    odbcGalileo.Close()
                Catch
                    DgvArchivio.Rows(ii).Cells(0).Value = "DOCUMENTO NON PROCESSATO"
                    DgvArchivio.Rows(ii).Cells(3).Value = "DOCUMENTO NON PROCESSATO"
                    DgvArchivio.Rows(ii).Cells(0).Style.BackColor = Color.Red
                    DgvArchivio.Rows(ii).Cells(1).Style.BackColor = Color.Red
                    DgvArchivio.Rows(ii).Cells(2).Style.BackColor = Color.Red
                    DgvArchivio.Rows(ii).Cells(3).Style.BackColor = Color.Red
                    GoTo 1
                End Try

                For Each rowCLI In tabCLI.Rows
                    If Not IsDBNull(rowCLI("DtDDT")) Then
                        tmpdata = rowCLI("DtDDT")
                        tmpdata = tmpdata.Substring(0, 4) & "/" & tmpdata.Substring(4, 2) & "/" & tmpdata.Substring(6, 2)
                        DgvArchivio.Rows(ii).Cells(3).Value = tmpdata
                    End If
                Next

                If DgvArchivio.Rows(ii).Cells(3).Value = "" Then

                    DgvArchivio.Rows(ii).Cells(0).Style.BackColor = Color.Red
                    DgvArchivio.Rows(ii).Cells(1).Style.BackColor = Color.Red
                    DgvArchivio.Rows(ii).Cells(2).Style.BackColor = Color.Red
                    DgvArchivio.Rows(ii).Cells(3).Style.BackColor = Color.Red

                    'MessageBox.Show("Documento non trovato!", "Errore!")
                Else

                    cnDB.ConnectionString = myConnectionString

                    ' CONTROLLO CHE IL DOCUMENTO IN INSERIMENTO NON SIA GIA' PRESENTE IN ARCHIVIO
                    query = "SELECT * FROM DOCUMENTI WHERE NUMDOCUMENTO=" & numdoc

                    myDAccess.SelectCommand = New OleDbCommand(query, cnDB)

                    Dim tmpTabDocumenti As New DataTable

                    cnDB.Open()

                    If myDAccess.Fill(tmpTabDocumenti) = 0 Then
                        query = "INSERT INTO DOCUMENTI VALUES ('" & tmpdata & "'," & numdoc & "," & _
                                _flagVettore & "," & _flagCliente & ")"
                        myDAccess.InsertCommand = New OleDbCommand(query, cnDB)
                        myDAccess.InsertCommand.ExecuteNonQuery()
                        cnDB.Close()

                        DgvArchivio.Rows(ii).Cells(0).Style.BackColor = Color.LightGreen
                        DgvArchivio.Rows(ii).Cells(0).ReadOnly = True

                        DgvArchivio.Rows(ii).Cells(1).Style.BackColor = Color.LightGreen
                        DgvArchivio.Rows(ii).Cells(2).Style.BackColor = Color.LightGreen
                        DgvArchivio.Rows(ii).Cells(3).Style.BackColor = Color.LightGreen
                        maxRiga += 1
                    Else
                        query = "UPDATE DOCUMENTI SET FIRMATAVETTORE=" & _flagVettore & ", FIRMATACLIENTE =" & _flagCliente & "" & _
                                " WHERE NUMDOCUMENTO=" & numdoc
                        myDAccess.UpdateCommand = New OleDbCommand(query, cnDB)
                        myDAccess.UpdateCommand.ExecuteNonQuery()
                        cnDB.Close()

                        DgvArchivio.Rows(ii).Cells(0).Style.BackColor = Color.LightGreen
                        DgvArchivio.Rows(ii).Cells(1).Style.BackColor = Color.LightGreen
                        DgvArchivio.Rows(ii).Cells(2).Style.BackColor = Color.LightGreen
                        DgvArchivio.Rows(ii).Cells(3).Style.BackColor = Color.LightGreen

                        'MessageBox.Show("Documento già presente aggiornato.")
                    End If
                End If
            End If
1:
        Next

    End Function

    Private Sub BtnControlla_Click_1(sender As System.Object, e As System.EventArgs) Handles BtnControlla.Click

        DgvDati.Rows.Clear()

        btnStampa.Enabled = False

        Dim NrDDTprecedente As Integer = 0
        Dim DataPrima As String = ""

        Cursor = Cursors.WaitCursor

        Mese = CmbMese.SelectedItem
        Anno = CmbAnno.SelectedItem
        AnnoDopo = Anno
        MeseDopo = Mese + 1

        If MeseDopo = "13" Then
            MeseDopo = "01"
            AnnoDopo = Anno + 1
        ElseIf CInt(MeseDopo) < 10 Then
            MeseDopo = "0" & MeseDopo
        End If

        DataI = Anno & Mese & "01"
        DataF = AnnoDopo & MeseDopo & "01"
        DataInizioAnno = Anno & "01" & "01"

        DgvDati.Rows.Clear()

        getBollettario(DataI, DataF)

        For Each rowBOL In tabBOL.Rows
            NrDDTprecedente = 0
            DataPrima = ""

            ' Trova l'ultimo ddt del mese prima
            query = "select MAG80DAT.FTMOV00F.NRBOFM AS NrDDT, " &
                    "MAG80DAT.FTMOV00F.DTBOFM as DtDDT " &
                    "from MAG80DAT.FTMOV00F " &
                    "where MAG80DAT.FTMOV00F.CDBOFM = " & "'" & rowBOL("BOLLETTARIO") & "'" &
                    " and MAG80DAT.FTMOV00F.DTBOFM >= " & "" & DataI & " and MAG80DAT.FTMOV00F.TDOCFM='B'" &
                    " UNION ALL " &
                    "select MAG80DAT.FTBKM00F.NRBOFM AS NrDDT, " &
                    "MAG80DAT.FTBKM00F.DTBOFM as DtDDT " &
                    "from MAG80DAT.FTBKM00F " &
                    "where MAG80DAT.FTBKM00F.CDBOFM = " & "'" & rowBOL("BOLLETTARIO") & "'" &
                    " and MAG80DAT.FTBKM00F.DTBOFM >= " & "" & DataI & " and MAG80DAT.FTBKM00F.TDOCFM='B'" &
                    " ORDER BY NrDDT ASC FETCH FIRST 1 ROW ONLY"

            myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
            tabDDT.Clear()
            myDA.Fill(tabDDT)

            For Each rowDDT In tabDDT.Rows
                NrDDTprecedente = rowDDT("NrDDT")
                DataPrima = rowDDT("DtDDT")
            Next

            query = "select MAG80DAT.FTMOV00F.NRBOFM AS NrDDT, " &
                    "MAG80DAT.FTMOV00F.DTBOFM as DtDDT " &
                    "from MAG80DAT.FTMOV00F " &
                    "where MAG80DAT.FTMOV00F.CDBOFM = " & "'" & rowBOL("BOLLETTARIO") & "'" &
                    " and MAG80DAT.FTMOV00F.DTBOFM >= " & "" & DataI & "" &
                    " and MAG80DAT.FTMOV00F.DTBOFM < " & "" & DataF & "" &
                    " and MAG80DAT.FTMOV00F.TDOCFM='B' AND MAG80DAT.FTMOV00F.NRBOFM<>" & NrDDTprecedente & "" &
                    " UNION ALL " &
                    "select MAG80DAT.FTBKM00F.NRBOFM AS NrDDT, " &
                    "MAG80DAT.FTBKM00F.DTBOFM as DtDDT " &
                    "from MAG80DAT.FTBKM00F " &
                    "where MAG80DAT.FTBKM00F.CDBOFM = " & "'" & rowBOL("BOLLETTARIO") & "'" &
                    " and MAG80DAT.FTBKM00F.DTBOFM >= " & "" & DataI & "" &
                    " and MAG80DAT.FTBKM00F.DTBOFM < " & "" & DataF & "" &
                    " and MAG80DAT.FTBKM00F.TDOCFM='B' AND MAG80DAT.FTBKM00F.NRBOFM<>" & NrDDTprecedente & "" &
                    " ORDER BY NrDDT ASC"

            myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
            tabDDT.Clear()
            myDA.Fill(tabDDT)

            For Each rowDDT In tabDDT.Rows

                If NrDDTprecedente = "28134" Then
                    Dim AA As Integer = 0
                    AA = 1
                End If
                If NrDDTprecedente > 0 Then

                    If (NrDDTprecedente + 1) <> rowDDT("NrDDT") Or NrDDTprecedente = rowDDT("NrDDT") Then

                        DgvDati.Rows.Add(New String() {rowBOL("BOLLETTARIO"), NrDDTprecedente.ToString, DataPrima})

                        If NrDDTprecedente = rowDDT("NrDDT") Then
                            DgvDati.Rows(DgvDati.Rows.Count - 1).Cells(0).Style.BackColor = Color.Red
                            DgvDati.Rows(DgvDati.Rows.Count - 1).Cells(1).Style.BackColor = Color.Red
                            DgvDati.Rows(DgvDati.Rows.Count - 1).Cells(2).Style.BackColor = Color.Red
                        Else
                            DgvDati.Rows(DgvDati.Rows.Count - 1).Cells(0).Style.BackColor = Color.Yellow
                            DgvDati.Rows(DgvDati.Rows.Count - 1).Cells(1).Style.BackColor = Color.Yellow
                            DgvDati.Rows(DgvDati.Rows.Count - 1).Cells(2).Style.BackColor = Color.Yellow
                        End If
                    End If
                    NrDDTprecedente = rowDDT("NrDDT")
                    DataPrima = rowDDT("DtDDT")
                End If
            Next
        Next
        Cursor = Cursors.Default
        If DgvDati.Rows.Count = 0 Then
            MessageBox.Show("NESSUNA ANOMALIA RISCONTRATA")
        Else
            MessageBox.Show(DgvDati.Rows.Count.ToString & " documenti presentano errori. " & vbCr & "Verranno mostrati gli ultimi documenti ritenuti corretti.", "Attenzione!")
        End If
        dtpStampa.Enabled = True
        btnStampa.Enabled = True
        odbcGalileo.Close()
    End Sub

    Private Sub btnStampa_Click(sender As System.Object, e As System.EventArgs) Handles btnStampa.Click

        If DgvDati.Rows.Count <> 0 Then
            Dim result As Integer = MessageBox.Show("Sono presenti errori sulla sequenza, procedo comunque con la stampa?", "Attenzione", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                MessageBox.Show("Stampa interrotta.")
                Exit Sub
            End If
        End If

        Cursor = Cursors.WaitCursor

        Dim DataStampa As String = ""

        DataStampa = dtpStampa.Value.Year.ToString

        If dtpStampa.Value.Month.ToString.Length > 1 Then
            DataStampa = DataStampa & "/" & dtpStampa.Value.Month.ToString
        Else
            DataStampa = DataStampa & "/0" & dtpStampa.Value.Month.ToString
        End If

        If dtpStampa.Value.Day.ToString.Length > 1 Then
            DataStampa = DataStampa & "/" & dtpStampa.Value.Day.ToString
        Else
            DataStampa = DataStampa & "/0" & dtpStampa.Value.Day.ToString
        End If

        query = "select MAG80DAT.FTMOV00F.NRBOFM AS NrBolla," & _
                " MAG80DAT.FTMOV00F.NRDFFM AS NrDocInterno, " & _
                " MAG80DAT.FTMOV00F.CDBOFM AS Bollettario," & _
                " MAG80DAT.FTMOV00F.CDCFFM  AS CodCliente, " & _
                " trim(MAG80DAT.CGANA01J.DSCOCP) AS RagSCliente " & _
                " from MAG80DAT.FTMOV00F " & _
                " LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTMOV00F.CDCFFM = MAG80DAT.CGANA01J.CONTCA " & _
                " where MAG80DAT.FTMOV00F.DTBOFM = " & "'" & DataStampa.Replace("/", "") & "'" & _
                " and MAG80DAT.FTMOV00F.TDOCFM='B'" & _
                " UNION ALL " & _
                " select MAG80DAT.FTBKM00F.NRBOFM AS NrBolla," & _
                " MAG80DAT.FTBKM00F.NRDFFM AS NrDocInterno, " & _
                " MAG80DAT.FTBKM00F.CDBOFM as Bollettario," & _
                " MAG80DAT.FTBKM00F.CDCFFM  AS CodCliente, " & _
                " trim(MAG80DAT.CGANA01J.DSCOCP) AS RagSCliente " & _
                " from MAG80DAT.FTBKM00F " & _
                " LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTBKM00F.CDCFFM = MAG80DAT.CGANA01J.CONTCA " & _
                " where MAG80DAT.FTBKM00F.DTBOFM = " & "'" & DataStampa.Replace("/", "") & "'" & _
                " and MAG80DAT.FTBKM00F.TDOCFM='B'" & _
                " ORDER BY Bollettario,NrBolla,NrDocInterno"


        odbcGalileo.ConnectionString = myConnectionStringGalileo
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabStampa.Clear()
        odbcGalileo.Open()
        Dim ii As Integer = myDA.Fill(tabStampa)
        odbcGalileo.Close()

        If ii > 0 Then


            Dim xlsAppl As New Microsoft.Office.Interop.Excel.Application


            xlsAppl.Visible = False

            With xlsAppl
                .Workbooks.Open(Filename:="\\SERVER\Principale\MAGIC PACK\PROCEDURE\APPLICAZIONI\ControlliFatturazione\Report\stampaBolleEmesse.xls")

                Dim idRiga As Integer = 2
                Dim idColonnaMax As Integer = 4

                Dim minDoc As String = ""
                Dim maxDoc As String = ""
                Dim tmpBollettario As String = ""
                Dim idRigaRiassuntivo As Integer = 36
                Dim idColonnaMaxRiassuntivo As Integer = 7
                Dim cambioBollettario As Boolean = False

                For Each rowStampa In tabStampa.Rows

                    If idRiga = 75 And idColonnaMax = 8 Then
                        ' fine documento, stampo e ricomincio
                        .Cells(35, 1) = Mid(DataStampa, 9, 2) & "/" & Mid(DataStampa, 6, 2)
                        .Cells(35, 8) = Mid(DataStampa, 1, 4)


                        If tmpBollettario = "" Then tmpBollettario = "001"

                        '.Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo - 3) = tmpBollettario
                        '.Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo - 1) = minDoc
                        '.Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo) = maxDoc
                        'idRigaRiassuntivo += 1

                        'tmpBollettario = ""
                        'minDoc = ""
                        'maxDoc = ""
                        .ActiveWindow.SelectedSheets.PrintOut(Copies:=1)                ' --- stampa -----

                        .Range("B2:H31").ClearContents()
                        .Range("B2:H31").Font.Bold = False

                        .Range("B45:I74").ClearContents()
                        .Range("B45:I74").Font.Bold = False

                        '.Range("D36:G41").ClearContents()

                        idRiga = 2
                        idColonnaMax = 4
                        'idRigaRiassuntivo = 36
                    End If

                    If tmpBollettario = "" Then
                        ' setto un bollettario
                        tmpBollettario = rowStampa("Bollettario")
                    End If

                    If rowStampa("Bollettario") = tmpBollettario And minDoc = "" Then
                        ' se non ho un doc minimo, setto il primo che trovo come minimo
                        minDoc = rowStampa("NrBolla")
                    ElseIf rowStampa("Bollettario") = tmpBollettario Then
                        ' tutti gli altri documenti sono considerati il massimo, a parità di bollettario
                        maxDoc = rowStampa("NrBolla")
                    ElseIf rowStampa("Bollettario") <> tmpBollettario Then
                        ' se il bollettario cambia, scrivo nell'xls e poi svuoto le variabili
                        If maxDoc = "" Then maxDoc = minDoc

                        If tmpBollettario = "" Then tmpBollettario = "001"

                        .Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo - 3) = tmpBollettario
                        .Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo - 1) = minDoc
                        .Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo) = maxDoc

                        idRigaRiassuntivo += 1

                        tmpBollettario = rowStampa("Bollettario")
                        minDoc = rowStampa("NrBolla")
                        maxDoc = ""
                        cambioBollettario = True
                    End If

                    If idRiga = 32 Then
                        idRiga = 45
                    End If

                    If idRiga = 75 Then
                        idRiga = 2
                        idColonnaMax = 8
                    End If

                    If tmpBollettario = "" Then tmpBollettario = "001"

                    .Cells(idRiga, idColonnaMax - 2) = tmpBollettario & " - " & rowStampa("NrBolla")
                    .Cells(idRiga, idColonnaMax - 1) = rowStampa("NrDocInterno")
                    .Cells(idRiga, idColonnaMax) = rowStampa("CodCliente") & " - " & rowStampa("RagSCliente")

                    If cambioBollettario Then
                        ' se ho appena settato un bollettario, marco in grassetto le celle indicate
                        .Cells(idRiga, idColonnaMax - 2).Font.Bold = True
                        .Cells(idRiga, idColonnaMax - 1).Font.Bold = True
                        .Cells(idRiga, idColonnaMax).Font.Bold = True
                        cambioBollettario = False
                    End If

                    idRiga += 1

                    If tmpBollettario = "001" Then tmpBollettario = ""

                Next

                .Cells(35, 1) = Mid(DataStampa, 9, 2) & "/" & Mid(DataStampa, 6, 2)
                .Cells(35, 8) = Mid(DataStampa, 1, 4)


                If tmpBollettario = "" Then tmpBollettario = "001"

                .Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo - 3) = tmpBollettario
                .Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo - 1) = minDoc
                .Cells(idRigaRiassuntivo, idColonnaMaxRiassuntivo) = maxDoc

                .Range("D35:G41").Font.Color = 0

                .ActiveWindow.SelectedSheets.PrintOut(Copies:=1)

                .Range("B2:H31").ClearContents()
                .Range("B2:H31").Font.Bold = False

                .Range("B45:I74").ClearContents()
                .Range("B45:I74").Font.Bold = False

                .Range("D36:G41").ClearContents()
                .Range("D36:G41").Font.Bold = False

                xlsAppl.DisplayAlerts = False

                .ActiveWorkbook.Close()

                xlsAppl.Quit()

                releaseObject(xlsAppl)
            End With

            Cursor = Cursors.Default
            MessageBox.Show("Stampa completata.")
        Else
            Cursor = Cursors.Default
            MessageBox.Show("Nessun documento da stampare.")
        End If


    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub BtnCheckMov1_Click(sender As System.Object, e As System.EventArgs) Handles BtnCheckMov1.Click
        Cursor = Cursors.WaitCursor
        pnlPersFiltri.Visible = False

        CheckMovimenti1()
        Cursor = Cursors.Default
        If DgvCheckMovimenti1.Rows.Count = 0 Then
            MessageBox.Show("NESSUNA ANOMALIA RISCONTRATA")
        End If

    End Sub

    Private Sub btnPersonalizza_Click(sender As System.Object, e As System.EventArgs) Handles btnPersonalizza.Click
        LoadCmbDestinatari()
        LoadPersonalizzaIva()
    End Sub

    Private Sub btnInserisciPersIVA_Click(sender As System.Object, e As System.EventArgs) Handles btnInserisciPersIVA.Click
        Dim Destinatario As String = ""
        Dim Iva As String = ""

        Destinatario = cmbCliente.SelectedItem
        ItemDgvPersIva(Destinatario, Iva)
    End Sub

    Public Sub tabPage9_click(ByVal sender As Object, ByVal e As EventArgs) Handles TabPage9.Click
        pnlPersFiltri.Visible = False
    End Sub
    Private Sub DgvPersiva_cellendedit(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPersIva.CellEndEdit
        Dim Destinatario As String
        Dim RagSoc As String
        Dim Iva As String

        Destinatario = Trim(dgvPersIva.Rows(e.RowIndex).Cells(0).Value)
        RagSoc = Trim(dgvPersIva.Rows(e.RowIndex).Cells(1).Value)
        Iva = Trim(dgvPersIva.Rows(e.RowIndex).Cells(2).Value)

        ManagePersonalizzaIva(Destinatario, RagSoc, Iva)
    End Sub
    Private Sub DgvCheckMovimenti1_CellContentDoubleClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvCheckMovimenti1.CellContentDoubleClick
        Dim Destinatario As String
        Dim RagSoc As String
        Dim Iva As String
        Dim Cod As String

        Destinatario = DgvCheckMovimenti1.Rows(e.RowIndex).Cells(1).Value
        If Destinatario > "" Then
            Iva = DgvCheckMovimenti1.Rows(e.RowIndex).Cells(2).Value
            Cod = Mid(Destinatario, 1, Destinatario.IndexOf(" "))
            RagSoc = Mid(Destinatario, Destinatario.IndexOf("-") + 3)

            LoadCmbDestinatari()
            LoadPersonalizzaIva()
            ItemDgvPersIva(Destinatario, Iva)
            ManagePersonalizzaIva(Cod, RagSoc, Iva)
        End If
    End Sub
    Private Sub LoadCmbDestinatari()
        If cmbCliente.Items.Count = 0 Then
            odbcGalileo.ConnectionString = myConnectionStringGalileo

            query = "select " & _
                    "RTrim(MAG80DAT.CGANA01J.CONTCA) AS CodDes, " & _
                    "RTrim(MAG80DAT.CGANA01J.DSCOCP) AS RagSoc  " & _
                    "from MAG80DAT.CGANA01J " & _
                    "where MAG80DAT.CGANA01J.CLFOCP='D' " & _
                    "order by 1"

            myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
            tabCLI.Clear()

            odbcGalileo.Open()
            myDA.Fill(tabCLI)
            odbcGalileo.Close()
            cmbCliente.Items.Clear()
            For Each rowCLI In tabCLI.Rows
                cmbCliente.Items.Add(rowCLI("CodDes") & " - " & Trim(rowCLI("RagSoc")))
            Next
        End If
    End Sub
    Private Sub LoadPersonalizzaIva()
        Dim Igdv As Integer
        cnDB.ConnectionString = myConnectionString

        query = "SELECT * FROM PERSONALIZZA_CONTROLLO_IVA ORDER BY DESTINATARIO"

        cnDB.Open()
        myDAccess.SelectCommand = New OleDbCommand(query, cnDB)
        tabIVA.Clear()
        myDAccess.Fill(tabIVA)
        cnDB.Close()

        dgvPersIva.Rows.Clear()
        Igdv = 0
        For Each rowIVA In tabIVA.Rows
            dgvPersIva.Rows.Add(New String() {rowIVA("DESTINATARIO"), Trim(rowIVA("RAGSOC")), rowIVA("IVA")})
        Next
        pnlPersFiltri.Visible = True
    End Sub

    Private Sub ItemDgvPersIva(ByVal Destinatario As String, ByVal Iva As String)
        Dim CodCLI As String
        Dim Ragsoc As String
        Dim Ncli As Integer

        If Destinatario > "" Then
            CodCLI = Mid(Destinatario, 1, Destinatario.IndexOf(" "))
            Ragsoc = Mid(Destinatario, Destinatario.IndexOf("-") + 3)

            dgvPersIva.Rows(0).Cells(0).Selected = False

            Ncli = 0
            Do While Ncli < dgvPersIva.Rows.Count AndAlso CodCLI <> dgvPersIva.Rows(Ncli).Cells(0).Value
                Ncli = Ncli + 1
            Loop
            If Ncli = dgvPersIva.Rows.Count Then
                dgvPersIva.Rows.Insert(0, New String() {CodCLI, Ragsoc, Iva})
                dgvPersIva.Rows(0).Cells(2).Selected = True
                dgvPersIva.Rows(1).Cells(2).Selected = False
                dgvPersIva.Rows(1).Cells(1).Selected = False
                dgvPersIva.Rows(1).Cells(0).Selected = False
            Else
                dgvPersIva.FirstDisplayedScrollingRowIndex = Ncli
                dgvPersIva.Rows(Ncli).Cells(2).Selected = True
            End If
        End If
    End Sub
    Private Sub ManagePersonalizzaIva(ByVal Destinatario As String, ByVal RagSoc As String, ByVal Iva As String)
        Dim IvaPrecedente As String = ""
        Dim Niva As Integer = 1

        If Iva > "" Then
            query = "SELECT trim(substring(MAG80DAT.SMTAB00F.XCODTB,1,3)) AS CodIva, " & _
                    "trim(substring(MAG80DAT.SMTAB00F.XDATTB,9,20)) AS DesIva, " & _
                    "rtrim(substring(MAG80DAT.SMTAB00F.XDATTB,29,2)) AS Aliquota " & _
                    "FROM MAG80DAT.SMTAB00F " & _
                    "WHERE MAG80DAT.SMTAB00F.XTIPTB='01CI' " & _
                    "and trim(substring(MAG80DAT.SMTAB00F.XCODTB,1,3)) = '" & Iva & "'"

            odbcGalileo.ConnectionString = myConnectionStringGalileo
            odbcGalileo.Open()
            myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
            tabIVA.Clear()
            Niva = myDA.Fill(tabIVA)
            odbcGalileo.Close()
            If Niva = 0 Then
                MessageBox.Show("Codice IVA sconosciuto")
            End If
        End If

        If Niva > 0 Then
            cnDB.ConnectionString = myConnectionString
            cnDB.Open()

            If Iva = "" Then
                query = "DELETE FROM PERSONALIZZA_CONTROLLO_IVA where DESTINATARIO = '" & Destinatario & "'"
                myDAccess.DeleteCommand = New OleDbCommand(query, cnDB)
                myDAccess.DeleteCommand.ExecuteNonQuery()

            Else
                query = "SELECT DESTINATARIO, IVA FROM PERSONALIZZA_CONTROLLO_IVA where DESTINATARIO = '" & Destinatario & "'"
                myDAccess.SelectCommand = New OleDbCommand(query, cnDB)
                tabIVA.Clear()
                myDAccess.Fill(tabIVA)
                IvaPrecedente = ""
                For Each rowIVA In tabIVA.Rows
                    IvaPrecedente = rowIVA("IVA")
                Next

                If IvaPrecedente > "" Then
                    If IvaPrecedente <> Iva Then
                        query = "UPDATE PERSONALIZZA_CONTROLLO_IVA " & _
                                "SET IVA = '" & Iva & "' " & _
                                "WHERE DESTINATARIO = '" & Destinatario & "'"
                        myDAccess.UpdateCommand = New OleDbCommand(query, cnDB)
                        myDAccess.UpdateCommand.ExecuteNonQuery()
                    End If
                Else
                    query = "INSERT INTO PERSONALIZZA_CONTROLLO_IVA VALUES ('" & Destinatario & "','" & _
                            RagSoc & "','" & Iva & "')"
                    myDAccess.InsertCommand = New OleDbCommand(query, cnDB)
                    myDAccess.InsertCommand.ExecuteNonQuery()
                End If
            End If
            cnDB.Close()

            CheckMovimenti1()
        End If
    End Sub
    Private Sub CheckMovimenti1()
        Dim IvaCorretta As String = ""
        Dim IvaItalia As String = ""
        Dim Destinatario As String = ""
        Dim tmpIvaAttuale As String = "" ' aggiunto per evitare dbnull su ogni codiva

        Cursor = Cursors.WaitCursor

        DataI = dtpDaCheck.Value.Year.ToString
        If dtpDaCheck.Value.Month.ToString.Length > 1 Then
            DataI = DataI & dtpDaCheck.Value.Month.ToString
        Else
            DataI = DataI & "0" & dtpDaCheck.Value.Month.ToString
        End If
        If dtpDaCheck.Value.Day.ToString.Length > 1 Then
            DataI = DataI & dtpDaCheck.Value.Day.ToString
        Else
            DataI = DataI & "0" & dtpDaCheck.Value.Day.ToString
        End If

        DataF = dtpACheck.Value.Year.ToString
        If dtpACheck.Value.Month.ToString.Length > 1 Then
            DataF = DataF & dtpACheck.Value.Month.ToString
        Else
            DataF = DataF & "0" & dtpACheck.Value.Month.ToString
        End If
        If dtpACheck.Value.Day.ToString.Length > 1 Then
            DataF = DataF & dtpACheck.Value.Day.ToString
        Else
            DataF = DataF & "0" & dtpACheck.Value.Day.ToString
        End If

        query = "SELECT MAG80DAT.FTMOV00F.NRDFFM AS NUMERO_DOC, " & _
                "MAG80DAT.FTMOV00F.DTBOFM AS DATA_BOLLA, " & _
                "MAG80DAT.CGANA01J.CONTCA AS COD_DEST, " & _
                "trim(MAG80DAT.CGANA01J.DSCOCP) AS DESTINATARIO, " & _
                "trim(MAG80DAT.CGANA01J.NAZICA) AS NAZIONE, " & _
                "MAG80DAT.FTMOV01F.ALIVFM AS COD_IVA, " & _
                "MAG80DAT.FTMOV01F.CDARFM AS ARTICOLO " & _
                "FROM (MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTMOV00F.CSPEFM=MAG80DAT.CGANA01J.CONTCA) " & _
                "LEFT JOIN MAG80DAT.FTMOV01F ON MAG80DAT.FTMOV00F.NRDFFM=MAG80DAT.FTMOV01F.NRDFFM and MAG80DAT.FTMOV00F.TDOCFM=MAG80DAT.FTMOV01F.TDOCFM " & _
                "where MAG80DAT.FTMOV00F.CDCFFM='" & Happy & "' and MAG80DAT.CGANA01J.CLFOCP='D' " & _
                "and MAG80DAT.FTMOV00F.DTBOFM >= " & "'" & DataI & "' " & _
                "and MAG80DAT.FTMOV00F.DTBOFM <= " & "'" & DataF & "' " & _
                "and MAG80DAT.FTMOV00F.TDOCFM ='B' " & _
                "and MAG80DAT.FTMOV01F.CDARFM >' ' " & _
                "UNION ALL " & _
                "SELECT MAG80DAT.FTBKM00F.NRDFFM AS NUMERO_DOC, " & _
                "MAG80DAT.FTBKM00F.DTBOFM AS DATA_BOLLA, " & _
                "MAG80DAT.CGANA01J.CONTCA AS COD_DEST, " & _
                "trim(MAG80DAT.CGANA01J.DSCOCP) AS DESTINATARIO, " & _
                "trim(MAG80DAT.CGANA01J.NAZICA) AS NAZIONE, " & _
                "MAG80DAT.FTBKM01F.ALIVFM AS COD_IVA, " & _
                "MAG80DAT.FTBKM01F.CDARFM AS ARTICOLO " & _
                "FROM (MAG80DAT.FTBKM00F LEFT JOIN MAG80DAT.CGANA01J ON MAG80DAT.FTBKM00F.CSPEFM=MAG80DAT.CGANA01J.CONTCA) " & _
                "LEFT JOIN MAG80DAT.FTBKM01F ON MAG80DAT.FTBKM00F.NRDFFM=MAG80DAT.FTBKM01F.NRDFFM and MAG80DAT.FTBKM00F.TDOCFM=MAG80DAT.FTBKM01F.TDOCFM " & _
                "where MAG80DAT.FTBKM00F.CDCFFM='" & Happy & "' and MAG80DAT.CGANA01J.CLFOCP='D' " & _
                "and MAG80DAT.FTBKM00F.DTBOFM >= " & "'" & DataI & "' " & _
                "and MAG80DAT.FTBKM00F.DTBOFM <= " & "'" & DataF & "' " & _
                "and MAG80DAT.FTBKM00F.TDOCFM ='B' " & _
                "and MAG80DAT.FTBKM01F.CDARFM >' ' " & _
                "order by 1"

        '             "and MAG80DAT.FTBKM01F.TIMOFM='01'"

        odbcGalileo.ConnectionString = myConnectionStringGalileo
        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabCheckMovimenti.Clear()
        Dim i As Integer = myDA.Fill(tabCheckMovimenti)
        odbcGalileo.Close()

        cnDB.ConnectionString = myConnectionString
        query = "SELECT IVA " & _
                "FROM PERSONALIZZA_CONTROLLO_IVA " & _
                "where DESTINATARIO = '----------' "
        cnDB.Open()
        myDAccess.SelectCommand = New OleDbCommand(query, cnDB)
        tabIVA.Clear()
        myDAccess.Fill(tabIVA)

        For Each rowIVA In tabIVA.Rows
            IvaItalia = rowIVA("IVA")
        Next

        query = "SELECT DESTINATARIO, IVA " & _
        "FROM PERSONALIZZA_CONTROLLO_IVA " & _
        "where DESTINATARIO <> '----------' " & _
        "ORDER BY DESTINATARIO"
        myDAccess.SelectCommand = New OleDbCommand(query, cnDB)
        tabIVA.Clear()
        myDAccess.Fill(tabIVA)
        cnDB.Close()

        DgvCheckMovimenti1.Rows.Clear()

        For Each rowCheckMovimenti In tabCheckMovimenti.Rows

            If Not IsDBNull(rowCheckMovimenti("NUMERO_DOC")) Then
                If rowCheckMovimenti("NUMERO_DOC").ToString = "398482" Then
                    Dim x As String
                    x = "AddressOf"
                End If
            End If

            If Not IsDBNull(rowCheckMovimenti("COD_IVA")) Then
                tmpIvaAttuale = rowCheckMovimenti("COD_IVA")
            End If

            IvaCorretta = ""
            For Each rowIVA In tabIVA.Rows
                If Trim(rowIVA("DESTINATARIO")) = rowCheckMovimenti("COD_DEST") Then
                    IvaCorretta = rowIVA("IVA")
                End If
            Next

            If IvaCorretta > "" Then
                If tmpIvaAttuale = IvaCorretta Then
                    IvaCorretta = ""
                End If
            Else
                If rowCheckMovimenti("NAZIONE") = "" Then
                    IvaCorretta = IvaItalia
                    If tmpIvaAttuale = IvaCorretta Then
                        IvaCorretta = ""
                    End If
                End If
            End If

            If IvaCorretta > "" Then
                Destinatario = rowCheckMovimenti("COD_DEST") & " - " & rowCheckMovimenti("DESTINATARIO")
                DgvCheckMovimenti1.Rows.Add(New String() {rowCheckMovimenti("NUMERO_DOC"), Destinatario, rowCheckMovimenti("ARTICOLO"), tmpIvaAttuale, IvaCorretta})
            End If
        Next
        Cursor = Cursors.Default
    End Sub
    Private Sub rbtNumBolla_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtNumBolla.CheckedChanged
        Dim DatI As String
        Dim DatF As String

        If Not changeSrcMode Then
            changeSrcMode = True
            rbtNumDoc.Checked = False
            rbtNumBolla.Checked = True
            srcViaNumDoc = False
            changeSrcMode = False

            LblAnnoDoc.Visible = True
            LblBollettarioDoc.Visible = True
            CmbBollettarioDoc.Visible = True
            CmbAnnoDoc.Visible = True

            CmbAnnoDoc.SelectedIndex = 0
        End If
    End Sub
    Private Sub rbtNumDoc_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtNumDoc.CheckedChanged
        If Not changeSrcMode Then
            changeSrcMode = True
            rbtNumBolla.Checked = False
            rbtNumDoc.Checked = True
            srcViaNumDoc = True
            changeSrcMode = False

            LblAnnoDoc.Visible = False
            LblBollettarioDoc.Visible = False
            CmbBollettarioDoc.Visible = False
            CmbAnnoDoc.Visible = False
        End If
    End Sub
    Private Sub findNumDoc(_docDa As String, _docA As String)

        Bollettario = CmbBollettarioDoc.SelectedItem

        query = "SELECT MAG80DAT.FTMOV00F.NRDFFM AS NrDOC " & _
                "from MAG80DAT.FTMOV00F " & _
                "where MAG80DAT.FTMOV00F.NRBOFM >= " & "'" & _docDa & "' " & _
                "and MAG80DAT.FTMOV00F.NRBOFM <= " & "'" & _docA & "' " & _
                "and MAG80DAT.FTMOV00F.DTBOFM >= " & " '" & DataI & "' " & _
                "and MAG80DAT.FTMOV00F.DTBOFM <= " & "'" & DataF & "' " & _
                "and MAG80DAT.FTMOV00F.TDOCFM='B' " & _
                "and MAG80DAT.FTMOV00F.CDBOFM = " & "'" & Bollettario & "' " & _
                "UNION ALL " & _
                "SELECT MAG80DAT.FTBKM00F.NRDFFM AS NrDOC " & _
                "from MAG80DAT.FTBKM00F " & _
                "where MAG80DAT.FTBKM00F.NRBOFM >= " & "'" & _docDa & "' " & _
                "and MAG80DAT.FTBKM00F.NRBOFM <= " & "'" & _docA & "' " & _
                "and MAG80DAT.FTBKM00F.DTBOFM >= " & " '" & DataI & "' " & _
                "and MAG80DAT.FTBKM00F.DTBOFM <= " & "'" & DataF & "' " & _
                "and MAG80DAT.FTBKM00F.TDOCFM='B' " & _
                "and MAG80DAT.FTBKM00F.CDBOFM = " & "'" & Bollettario & "' "

        odbcGalileo.ConnectionString = myConnectionStringGalileo
        odbcGalileo.Open()
        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        TMPtabBol.Clear()
        myDA.Fill(TMPtabBol)
        odbcGalileo.Close()
    End Sub

    Private Sub CmbBOLLETTARIO_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CmbBOLLETTARIO.SelectedIndexChanged

        Bollettario = CmbBOLLETTARIO.SelectedItem

        CompilaDgvCheckPresenzeDoc1()

    End Sub

    Private Sub CmbSTATO_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CmbSTATO.SelectedIndexChanged

        Stato = CmbSTATO.SelectedItem
        CompilaDgvCheckPresenzeDoc1()

    End Sub
    Private Sub CompilaDgvCheckPresenzeDoc1()
        Dim II As Integer = 0
        Dim Boll As String = ""
        Dim Stat As String = ""
        Dim Carica As Boolean
        Dim Ig As Integer

        If Stato = " " Then
            Stato = "Tutti"
        End If
        If Bollettario = " " Then
            Bollettario = "Tutti"
        End If

        dgvCheckPresenzaDoc1.Rows.Clear()
        Ig = 0
        If Bollettario = "Tutti" And Stato = "Tutti" Then
            dgvCheckPresenzaDoc.Visible = True
            dgvCheckPresenzaDoc1.Visible = False
        Else
            For II = 0 To dgvCheckPresenzaDoc.Rows.Count - 1
                Boll = dgvCheckPresenzaDoc.Rows(II).Cells(1).Value
                Stat = dgvCheckPresenzaDoc.Rows(II).Cells(2).Value
                Carica = False

                If Bollettario <> "Tutti" And Stato <> "Tutti" Then
                    If Boll = Bollettario And Stato = Stat Then
                        Carica = True
                    End If
                ElseIf Bollettario = "Tutti" Then
                    If Stato = Stat Then
                        Carica = True
                    End If
                Else
                    If Bollettario = Boll Then
                        Carica = True
                    End If
                End If

                If Carica = True Then
                    dgvCheckPresenzaDoc1.Rows.Add(New String() {"", "", "", "", "", "", ""})

                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(0).Value = dgvCheckPresenzaDoc.Rows(II).Cells(0).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(1).Value = dgvCheckPresenzaDoc.Rows(II).Cells(1).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(2).Value = dgvCheckPresenzaDoc.Rows(II).Cells(2).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(3).Value = dgvCheckPresenzaDoc.Rows(II).Cells(3).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(4).Value = dgvCheckPresenzaDoc.Rows(II).Cells(4).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(5).Value = dgvCheckPresenzaDoc.Rows(II).Cells(5).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(6).Value = dgvCheckPresenzaDoc.Rows(II).Cells(6).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(7).Value = dgvCheckPresenzaDoc.Rows(II).Cells(7).Value
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(8).Value = dgvCheckPresenzaDoc.Rows(II).Cells(8).Value

                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(0).Style = dgvCheckPresenzaDoc.Rows(II).Cells(0).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(1).Style = dgvCheckPresenzaDoc.Rows(II).Cells(1).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(2).Style = dgvCheckPresenzaDoc.Rows(II).Cells(2).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(3).Style = dgvCheckPresenzaDoc.Rows(II).Cells(3).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(4).Style = dgvCheckPresenzaDoc.Rows(II).Cells(4).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(5).Style = dgvCheckPresenzaDoc.Rows(II).Cells(5).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(6).Style = dgvCheckPresenzaDoc.Rows(II).Cells(6).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(7).Style = dgvCheckPresenzaDoc.Rows(II).Cells(7).Style
                    dgvCheckPresenzaDoc1.Rows(Ig).Cells(8).Style = dgvCheckPresenzaDoc.Rows(II).Cells(8).Style

                    Ig = Ig + 1
                End If
            Next

            dgvCheckPresenzaDoc.Visible = False
            dgvCheckPresenzaDoc1.Visible = True
        End If
    End Sub
    Private Sub Btn369_Click(sender As System.Object, e As System.EventArgs) Handles Btn369.Click

        Dim Cliente As String = ""
        Dim Articolo As String = ""
        Dim DatDoc As String

        Cursor = Cursors.WaitCursor

        Dgv369.Rows.Clear()

        DataI = DtpDa369.Value.Year.ToString

        If DtpDa369.Value.Month.ToString.Length > 1 Then
            DataI = DataI & DtpDa369.Value.Month.ToString
        Else
            DataI = DataI & "0" & DtpDa369.Value.Month.ToString
        End If

        If DtpDa369.Value.Day.ToString.Length > 1 Then
            DataI = DataI & DtpDa369.Value.Day.ToString
        Else
            DataI = DataI & "0" & DtpDa369.Value.Day.ToString
        End If

        DataF = DtpA369.Value.Year.ToString
        If DtpA369.Value.Month.ToString.Length > 1 Then
            DataF = DataF & DtpA369.Value.Month.ToString
        Else
            DataF = DataF & "0" & DtpA369.Value.Month.ToString
        End If
        If DtpA369.Value.Day.ToString.Length > 1 Then
            DataF = DataF & DtpA369.Value.Day.ToString
        Else
            DataF = DataF & "0" & DtpA369.Value.Day.ToString
        End If

        query = "select MAG80DAT.FTMOV00F.NRDFFM AS NrDoc, " &
                "MAG80DAT.FTMOV00F.DTBOFM AS DtDoc, " &
                "MAG80DAT.FTMOV00F.CDCFFM AS CLIENTE, " &
                "MAG80DAT.CGANA01J.DSCOCP AS RAGCLI, " &
                "TRIM(MAG80DAT.FTMOV01F.CDARFM) AS ARTICOLO, " &
                "MAG80DAT.FTMOV01F.PRZUFM AS PREZZO, " &
                "MAG80DAT.MGART00F.DSARMA AS DESART " &
                "from ((MAG80DAT.FTMOV00F LEFT JOIN MAG80DAT.CGANA01J ON " &
                "MAG80DAT.FTMOV00F.CDCFFM=MAG80DAT.CGANA01J.CONTCA) " &
                "LEFT JOIN MAG80DAT.FTMOV01F ON MAG80DAT.FTMOV00F.NRDFFM=MAG80DAT.FTMOV01F.NRDFFM and " &
                "MAG80DAT.FTMOV00F.TDOCFM=MAG80DAT.FTMOV01F.TDOCFM) " &
                "LEFT JOIN MAG80DAT.MGART00F ON MAG80DAT.FTMOV01F.CDARFM=MAG80DAT.MGART00F.CDARMA " &
                "where MAG80DAT.FTMOV00F.DTBOFM >= " & "'" & DataI & "' " &
                "and MAG80DAT.FTMOV00F.DTBOFM <= " & "'" & DataF & "' " &
                "and MAG80DAT.FTMOV00F.TDOCFM='B' " &
                "and MAG80DAT.FTMOV01F.PRZUFM IN (3,4,5,6,7,8,9) " &
                "ORDER BY NrDoc"

        myDA.SelectCommand = New OdbcCommand(query, odbcGalileo)
        tabDDT.Clear()
        odbcGalileo.Open()
        myDA.Fill(tabDDT)
        odbcGalileo.Close()

        For Each rowDDT In tabDDT.Rows
            Cliente = rowDDT("CLIENTE") & " - " & rowDDT("RAGCLI")
            Articolo = rowDDT("ARTICOLO") & " - " & rowDDT("DESART")
            DatDoc = rowDDT("dtDOC")
            DatDoc = DatDoc.Substring(6, 2) & "/" & DatDoc.Substring(4, 2) & "/" & DatDoc.Substring(0, 4)
            Dgv369.Rows.Add(New String() {rowDDT("NrDoc"), DatDoc, rowDDT("PREZZO"), Articolo, Cliente})
        Next
        Cursor = Cursors.Default
        If Dgv369.Rows.Count = 0 Then
            MessageBox.Show("NESSUNA ANOMALIA RISCONTRATA")
        End If
    End Sub

    Private Sub CmbAnnoDoc_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CmbAnnoDoc.SelectedIndexChanged
        Dim DatI As String
        Dim DatF As String

        Anno = CmbAnnoDoc.SelectedItem
        DatI = Anno + "01" + "01"
        DatF = Anno + "12" + "31"
        DataI = DatI
        DataF = DatF

        CmbBollettarioDoc.Items.Clear()
        getBollettario(DatI, DatF)

        For Each rowBOL In tabBOL.Rows
            CmbBollettarioDoc.Items.Add(rowBOL("BOLLETTARIO"))
        Next
        CmbBollettarioDoc.SelectedIndex = CmbBollettarioDoc.Items.Count - 1
    End Sub

    Private Sub btnCreaReport_Click(sender As System.Object, e As System.EventArgs) Handles btnCreaReport.Click

        Dim xlsAppl As New Microsoft.Office.Interop.Excel.Application
        Dim FileExcel As String = "\\SERVER\Principale\MAGIC PACK\PROCEDURE\APPLICAZIONI\ControlliFatturazione\Report\ReportArchivio.xls"
        Dim ProgrammaExcel As String = "\\SERVER\Principale\MAGIC PACK\PROCEDURE\APPLICAZIONI\ControlliFatturazione\Report\ReportArchivio_Master.xls"
        Dim RigaXLS As Integer = 2
        Dim Passo As Integer = 0

        xlsAppl.Visible = False

        With xlsAppl
            .Workbooks.Open(Filename:=ProgrammaExcel)
            If dgvCheckPresenzaDoc.Visible Then
                For i As Integer = 0 To dgvCheckPresenzaDoc.Rows.Count - 1

                    Dim tmpdata As String = Mid(dgvCheckPresenzaDoc.Rows(i).Cells("colDataBolla").Value.ToString, 1, 4) & "-"

                    tmpdata = tmpdata & Mid(dgvCheckPresenzaDoc.Rows(i).Cells("colDataBolla").Value.ToString, 5, 2) & "-"

                    tmpdata = tmpdata & Mid(dgvCheckPresenzaDoc.Rows(i).Cells("colDataBolla").Value.ToString, 7, 2)

                    .Cells(RigaXLS, 1) = dgvCheckPresenzaDoc.Rows(i).Cells("colNrBolla").Value.ToString
                    .Cells(RigaXLS, 2) = tmpdata
                    .Cells(RigaXLS, 3) = dgvCheckPresenzaDoc.Rows(i).Cells("colCliente").Value.ToString
                    .Cells(RigaXLS, 4) = dgvCheckPresenzaDoc.Rows(i).Cells("colDestinatario").Value.ToString
                    .Cells(RigaXLS, 5) = dgvCheckPresenzaDoc.Rows(i).Cells("colPrimoVettore").Value.ToString
                    .Cells(RigaXLS, 6) = dgvCheckPresenzaDoc.Rows(i).Cells("colSecondoVettore").Value.ToString

                    RigaXLS += 1

                Next
            ElseIf dgvCheckPresenzaDoc1.Visible Then
                For i As Integer = 0 To dgvCheckPresenzaDoc1.Rows.Count - 1

                    Dim tmpdata As String = Mid(dgvCheckPresenzaDoc1.Rows(i).Cells("colDataBolla_1").Value.ToString, 1, 4) & "-"

                    tmpdata = tmpdata & Mid(dgvCheckPresenzaDoc1.Rows(i).Cells("colDataBolla_1").Value.ToString, 5, 2) & "-"

                    tmpdata = tmpdata & Mid(dgvCheckPresenzaDoc1.Rows(i).Cells("colDataBolla_1").Value.ToString, 7, 2)

                    .Cells(RigaXLS, 1) = dgvCheckPresenzaDoc1.Rows(i).Cells("colNrBolla_1").Value.ToString
                    .Cells(RigaXLS, 2) = tmpdata
                    .Cells(RigaXLS, 3) = dgvCheckPresenzaDoc1.Rows(i).Cells("colCliente_1").Value.ToString
                    .Cells(RigaXLS, 4) = dgvCheckPresenzaDoc1.Rows(i).Cells("colDestinatario_1").Value.ToString
                    .Cells(RigaXLS, 5) = dgvCheckPresenzaDoc1.Rows(i).Cells("colPrimoVettore_1").Value.ToString
                    .Cells(RigaXLS, 6) = dgvCheckPresenzaDoc1.Rows(i).Cells("colSecondoVettore_1").Value.ToString


                    RigaXLS += 1

                Next
            End If

            .DisplayAlerts = False
            .ActiveWorkbook.SaveAs(Filename:=FileExcel)

            xlsAppl.Visible = True

        End With
    End Sub

    Private Sub txtDaDoc_Enter(sender As System.Object, e As System.EventArgs) Handles txtDaDoc.Enter
        txtDaDoc.Text = ""
    End Sub

    Private Sub txtADoc_Enter(sender As System.Object, e As System.EventArgs) Handles txtADoc.Enter
        txtADoc.Text = ""
    End Sub

    Private Sub btnAnnullaUltima_Click(sender As System.Object, e As System.EventArgs)

        If lastNrDoc <> "" Then

            Dim idRowToDelete As String = ""

            For ii As Integer = 0 To DgvArchivio.Rows.Count - 1

                If DgvArchivio.Rows(ii).Cells("DgvDocumento").Value <> Nothing AndAlso DgvArchivio.Rows(ii).Cells("DgvDocumento").Value.ToString = lastNrDoc Then
                    idRowToDelete = ii
                End If

            Next

            If idRowToDelete <> "" Then
                Dim result As Integer = MessageBox.Show("Procedo con l'eliminazione del documento N° " & lastNrDoc & " ?", "Attenzione", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    MessageBox.Show("Eliminazione interrotta.")
                    Exit Sub
                Else
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvDocumento").Value = ""
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvVETTORE").Value = 0
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvCliente1").Value = 0
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvDataDoc").Value = ""

                    DgvArchivio.Rows(idRowToDelete).Cells("DgvDocumento").Style.BackColor = Color.White
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvVETTORE").Style.BackColor = Color.White
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvCliente1").Style.BackColor = Color.White
                    DgvArchivio.Rows(idRowToDelete).Cells("DgvDataDoc").Style.BackColor = Color.White

                End If
            End If

            lastNrDoc = ""

            For ii As Integer = 0 To DgvArchivio.Rows.Count - 1

                If DgvArchivio.Rows(ii).Cells("DgvDocumento").Value = Nothing AndAlso DgvArchivio.Rows(ii + 1).Cells("DgvDocumento").Value = Nothing Then
                    lastNrDoc = DgvArchivio.Rows(ii).Cells("DgvDocumento").Value.ToString
                    Exit Sub
                End If

            Next
        Else
            MessageBox.Show("Operazione non disponibile.", "Attenzione")
        End If
    End Sub

    Private Sub DgvArchivio_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DgvArchivio.CellClick

        If e.ColumnIndex = 4 And e.RowIndex >= 0 Then

            If DgvArchivio.Rows(e.RowIndex).Cells(0).Value <> "" And DgvArchivio.Rows(e.RowIndex).Cells(0).Value <> "CANCELLATO" Then


                Dim DocToDelete As String = DgvArchivio.Rows(e.RowIndex).Cells(0).Value

                Dim result As Integer = MessageBox.Show("Procedo con l'eliminazione del documento N° " & DocToDelete & " ?", "Attenzione", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    MessageBox.Show("Eliminazione interrotta.")
                    Exit Sub
                Else
                    For ii As Integer = 0 To DgvArchivio.Rows.Count - 1

                        If DgvArchivio.Rows(ii).Cells("DgvDocumento").Value <> Nothing AndAlso DgvArchivio.Rows(ii).Cells("DgvDocumento").Value.ToString = DocToDelete Then

                            DgvArchivio.Rows(ii).Cells("DgvDocumento").Value = "CANCELLATO"
                            DgvArchivio.Rows(ii).Cells("DgvVETTORE").Value = 0
                            DgvArchivio.Rows(ii).Cells("DgvCliente1").Value = 0
                            DgvArchivio.Rows(ii).Cells("DgvDataDoc").Value = "CANCELLATO"

                            DgvArchivio.Rows(ii).Cells("DgvDocumento").Style.BackColor = Color.Gray
                            DgvArchivio.Rows(ii).Cells("DgvVETTORE").Style.BackColor = Color.Gray
                            DgvArchivio.Rows(ii).Cells("DgvCliente1").Style.BackColor = Color.Gray
                            DgvArchivio.Rows(ii).Cells("DgvDataDoc").Style.BackColor = Color.Gray

                        End If

                    Next

                    cnDB.Open()
                    query = "DELETE FROM DOCUMENTI WHERE NUMDOCUMENTO=" & DocToDelete
                    myDAccess.DeleteCommand = New OleDbCommand(query, cnDB)
                    myDAccess.DeleteCommand.ExecuteNonQuery()
                    cnDB.Close()

                    MessageBox.Show("Eliminazione completata.")

                End If

            End If
        End If

    End Sub

    Private Sub checkNumDoc_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles checkNumDoc.CheckedChanged
        pnlRicercaDoc.Enabled = checkNumDoc.Checked
    End Sub
End Class
