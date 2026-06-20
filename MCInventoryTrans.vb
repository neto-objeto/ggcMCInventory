'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
'
' Copyright 2006 and beyond
' All Rights Reserved
'
'     MC Inventory Transaction object
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Kalyptus [ 09/10/2016 10:11 am ]
'     Started creating this object based on ggcMCInventory-VB6 edition.
Imports MySql.Data.MySqlClient
Imports ADODB
Imports ggcAppDriver

Public Class MCInventoryTrans
    Public Const pxeMCAcceptance As String = "MCDA"
    Public Const pxeMCAcceptDelivery As String = "MCDl"
    Public Const pxeMCAcceptBackLoad As String = "MCAB"
    Public Const pxeMCDelivery As String = "MCDv"
    Public Const pxeMCBackLoad As String = "MCBB"
    Public Const pxeMCBackLoadTrucking As String = "MCBT"
    Public Const pxeMCPurchaseReturn As String = "MCPR"
    Public Const pxeMCSales As String = "MCSl"
    Public Const pxeMCRelease As String = "MCRl"
    Public Const pxeMCImpound As String = "MCIm"
    Public Const pxeMCSalesReturn As String = "MCSR"
    Public Const pxeMCSalesReplacement As String = "MCSp"
    Public Const pxeMCWarrantyRelease As String = "MCWR"
    Public Const pxeMCWarrantyPullOut As String = "MCWP"
    Public Const pxeMCAssume As String = "MCAs"
    Public Const pxeMCRegister As String = "MCRg"
    Public Const pxeMCTransfer As String = "MCTO"
    Public Const pxeMCTransferRenewal As String = "MCTR"
    Public Const pxeMCClearRecvd As String = "MCAC"
    Public Const pxeMCSalesInvoice As String = "MCIn"
    Public Const pxeMCAdjustment As String = "MCAd"

    'Actual part of code starts here
    Private Const pxeLgrNoPict As String = "000000"

    Private p_oApp As GRider
    Private p_oDTMstr As DataTable
    Private p_oDTDetl As DataTable

    Private p_sBranchCd As String
    Private p_dTransact As Date
    Private p_sSourceCd As String
    Private p_sSourceNo As String

    Private p_bInitTran As Boolean

    Public Property AppDriver() As ggcAppDriver.GRider
        Get
            Return p_oApp
        End Get
        Set(value As ggcAppDriver.GRider)
            p_oApp = value
        End Set
    End Property

    Public Property Branch() As String
        Get
            Return p_sBranchCd
        End Get
        Set(value As String)
            p_sBranchCd = value
        End Set
    End Property

    Public Property Detail(Row As Integer, ByVal Index As String) As Object
        Get
            If Not p_bInitTran Then Return ""
            Return p_oDTMstr(Row).Item(Index)
        End Get

        Set(value As Object)
            If Not p_bInitTran Then Exit Property
            If Row < 0 Then Exit Property

            If Row > p_oDTMstr.Rows.Count Then
                Exit Property
            ElseIf Row > p_oDTMstr.Rows.Count - 1 Then
                AddDetail()
            End If

            Select Case LCase(Index)
                Case "smcinvidx"
                    p_oDTMstr(Row).Item(Index) = value
                Case "nqtyinxxx", "nqtyoutxx"
                    p_oDTMstr(Row).Item(Index) = value
                Case "nrpoqtyin", "nrpoqtyot"
                    p_oDTMstr(Row).Item(Index) = value
            End Select
        End Set
    End Property

    Public ReadOnly Property ItemCount() As Long
        Get
            Return p_oDTMstr.Rows.Count
        End Get
    End Property

    Public Function InitTransaction(sSourceCd As String) As Boolean
        Select Case sSourceCd
            Case MCInventoryTrans.pxeMCAcceptance
            Case MCInventoryTrans.pxeMCAcceptDelivery
            Case MCInventoryTrans.pxeMCAcceptBackLoad
            Case MCInventoryTrans.pxeMCDelivery
            Case MCInventoryTrans.pxeMCBackLoad
            Case MCInventoryTrans.pxeMCBackLoadTrucking
            Case MCInventoryTrans.pxeMCPurchaseReturn
            Case MCInventoryTrans.pxeMCSales
            Case MCInventoryTrans.pxeMCRelease
            Case MCInventoryTrans.pxeMCImpound
            Case MCInventoryTrans.pxeMCSalesReturn
            Case MCInventoryTrans.pxeMCWarrantyRelease
            Case MCInventoryTrans.pxeMCWarrantyPullOut
            Case MCInventoryTrans.pxeMCAdjustment
            Case Else
                MsgBox("Invalid Transaction Source Detected!", vbCritical, "Warning")
                Return False
        End Select

        p_sSourceCd = sSourceCd

        CreateTransaction()
        Call AddDetail()

        p_bInitTran = True
        InitTransaction = True

    End Function

    Public Function SaveTransaction(sSourceNo As String, _
                             dTransact As Date) As Boolean
        Dim lnMstrRow As Long

        p_dTransact = dTransact
        p_sSourceNo = sSourceNo

        Try
            With p_oDTMstr
                lnMstrRow = 0
                Do While .Rows.Count > lnMstrRow
                    If .Rows(lnMstrRow).Item("sMCInvIDx") = "" Then Exit Do

                    If saveDetail(lnMstrRow) = False Then Return False

                    lnMstrRow = lnMstrRow + 1
                Loop
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        SaveTransaction = True
    End Function

    Public Function DeleteTransaction(sSourceNo As String, _
                               dTransact As Date) As Boolean
        p_dTransact = dTransact
        p_sSourceNo = sSourceNo

        Try
            ' now load transaction regardless of mode
            If LoadTransaction() = False Then Return False

            Dim lnRow As Integer

            lnRow = 0
            With p_oDTDetl
                Do While .Rows.Count > lnRow
                    If delDetail(.Rows(lnRow).Item("sMCInvIDx"), _
                                 .Rows(lnRow).Item("nQtyInxxx"), _
                                 .Rows(lnRow).Item("nQtyOutxx"), _
                                 .Rows(lnRow).Item("nRpoQtyIn"), _
                                 .Rows(lnRow).Item("nRpoQtyOt"), _
                                 .Rows(lnRow).Item("nLedgerNo")) = False Then Return False
                    lnRow = lnRow + 1
                Loop
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        DeleteTransaction = True
    End Function

    Private Function LoadTransaction() As Boolean
        'Test if the object was initialized
        If p_bInitTran = False Then
            MsgBox("Object is not initialized!!!" & vbCrLf & vbCrLf & _
                  "Please inform the SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning")
            Return False
        End If

        Try
            Dim lsSQL As String
            lsSQL = "SELECT" & _
                        "  a.sMCInvIDx" & _
                        ", a.nLedgerNo" & _
                        ", a.dTransact" & _
                        ", a.nQtyInxxx" & _
                        ", a.nQtyOutxx" & _
                        ", a.nRpoQtyIn" & _
                        ", a.nRpoQtyOt" & _
                        ", a.nQtyOnHnd" & _
                        ", a.nRpoOnHnd" & _
                        ", b.nLedgerNo xLedgerNo" & _
                     " FROM MC_Inventory_Ledger a" & _
                        ", MC_Inventory b"
            lsSQL = lsSQL & _
                     " WHERE a.sMCInvIDx = b.sMCInvIDx" & _
                        " AND a.sBranchCd = b.sBranchCd" & _
                        " AND a.sBranchCd = " & strParm(p_sBranchCd) & _
                        " AND a.sSourceNo = " & strParm(p_sSourceNo) & _
                        " AND a.sSourceCd = " & strParm(p_sSourceCd) & _
                     " ORDER BY a.sMCInvIDx"

            p_oDTDetl = p_oApp.ExecuteQuery(lsSQL)
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        LoadTransaction = True
    End Function

    Private Sub AddDetail()
        With p_oDTMstr
            .Rows.Add(.NewRow)
            .Rows(.Rows.Count - 1).Item("sMCInvIDx") = ""
            .Rows(.Rows.Count - 1).Item("nBegQtyxx") = 0
            .Rows(.Rows.Count - 1).Item("nRpoQtyxx") = 0
            .Rows(.Rows.Count - 1).Item("nQtyOnHnd") = 0
            .Rows(.Rows.Count - 1).Item("nRpoOnHnd") = 0
            .Rows(.Rows.Count - 1).Item("nQtyInxxx") = 0
            .Rows(.Rows.Count - 1).Item("nQtyOutxx") = 0
            .Rows(.Rows.Count - 1).Item("nRpoQtyIn") = 0
            .Rows(.Rows.Count - 1).Item("nRpoQtyOt") = 0
            .Rows(.Rows.Count - 1).Item("nPurPrice") = 0
        End With
    End Sub

    Private Function delDetail(lsMCInvIDx As String, _
                               lnQtyInxxx As Integer, _
                               lnQtyOutxx As Integer, _
                               lnRpoQtyIn As Integer, _
                               lnRpoQtyOt As Integer, _
                               lnLedgerNo As Integer) As Boolean
        Dim lsMasSQL As String, lsLgrSQL As String
        Dim lnRow As Long

        Try
            lsMasSQL = "UPDATE MC_Inventory SET" & _
                           "  nQtyOnHnd = nQtyOnHnd + " & lnQtyOutxx - lnQtyInxxx & _
                           ", nRpoOnHnd = nRpoOnHnd + " & lnRpoQtyOt - lnRpoQtyIn & _
                           ", nLedgerNo = " & strParm(Format(lnLedgerNo - 1, pxeLgrNoPict)) & _
                           ", dModified = " & dateParm(p_oApp.getSysDate) & _
                        " WHERE sMCInvIDx = " & strParm(lsMCInvIDx) & _
                           " AND sBranchCd = " & strParm(p_sBranchCd)

            lsLgrSQL = "DELETE FROM MC_Inventory_Ledger" & _
                        " WHERE sMCInvIDx = " & strParm(lsMCInvIDx) & _
                           " AND sSourceCd = " & strParm(p_sSourceCd) & _
                           " AND sSourceNo = " & strParm(p_sSourceNo)

            lnRow = p_oApp.Execute(lsMasSQL, "MC_Inventory", p_sBranchCd)
            If lnRow <= 0 Then
                MsgBox("Unable to Update Motorcycle Inventory!", vbCritical, "Warning")
                Return False
            End If

            lnRow = p_oApp.Execute(lsLgrSQL, "MC_Inventory_Ledger", p_sBranchCd)
            If lnRow <= 0 Then
                MsgBox("Unable to Update Motorcycle Inventory Transaction!", vbCritical, "Warning")
                Return False
            End If

            If lnLedgerNo - 1 > 1 Then
                If recalcOnHand(lsMCInvIDx, _
                      lnLedgerNo) = False Then Return False
            End If
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        delDetail = True
    End Function

    Private Function recalcOnHand(lsMCInvIDx As String, _
             ByVal lnLedgerNo As Long) As Boolean
        Dim loDta As DataTable
        Dim lsSQL As String, lnRow As Long
        Dim lnQtyOnHnd As Integer, lnRpoOnHnd As Integer

        Try
            lsSQL = "SELECT * FROM MC_Inventory_Ledger" & _
                     " WHERE sMCInvIDx = " & strParm(lsMCInvIDx) & _
                       " AND nLedgerNo >= " & strParm(Format(lnLedgerNo, pxeLgrNoPict)) & _
                       " AND sBranchCd = " & strParm(p_sBranchCd)

            loDta = p_oApp.ExecuteQuery(lsSQL)

            With loDta
                If loDta.Rows.Count > 0 Then
                    lnQtyOnHnd = .Rows(0).Item("nQtyOnHnd")
                    lnRpoOnHnd = .Rows(0).Item("nRpoOnHnd")
                    lnLedgerNo = lnLedgerNo + 1
                    lnRow = 1
                    Do While .Rows.Count > lnRow
                        lnQtyOnHnd = lnQtyOnHnd + .Rows(lnRow).Item("nQtyInxxx") - .Rows(lnRow).Item("nQtyOutxx")
                        lnRpoOnHnd = lnRpoOnHnd + .Rows(lnRow).Item("nRpoQtyIn") - .Rows(lnRow).Item("nRpoQtyOt")
                        lsSQL = "UPDATE MC_Inventory_Ledger SET" & _
                                    "  nQtyOnHnd = " & lnQtyOnHnd & _
                                    ", nRpoOnHnd = " & lnRpoOnHnd & _
                                    ", nLedgerNo = " & strParm(Format(lnLedgerNo, pxeLgrNoPict)) & _
                                    ", dModified = " & dateParm(p_oApp.getSysDate) & _
                                 " WHERE sMCInvIDx = " & strParm(lsMCInvIDx) & _
                                    " AND sBranchCd = " & strParm(p_sBranchCd) & _
                                    " AND nLedgerNo = " & strParm(.Rows(lnRow).Item("nLedgerNo"))

                        lnRow = p_oApp.Execute(lsSQL, "MC_Inventory_Ledger", p_sBranchCd)
                        If lnRow = 0 Then
                            MsgBox("Unable to Update MC Inventory Ledger!", vbCritical, "Warning")
                            Return False
                        End If

                        lnLedgerNo = lnLedgerNo + 1
                        lnRow = lnRow + 1
                    Loop
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        recalcOnHand = True
    End Function

    Private Sub CreateTransaction()
        p_oDTMstr = New DataTable
        p_oDTMstr.Columns.Add("sMCInvIDx", Type.GetType("System.String"))
        p_oDTMstr.Columns.Add("dTransact", Type.GetType("System.DateTime"))
        p_oDTMstr.Columns.Add("nBegQtyxx", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nRpoQtyxx", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nQtyOnHnd", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nRpoOnHnd", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nQtyInxxx", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nQtyOutxx", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nRpoQtyIn", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nRpoQtyOt", Type.GetType("System.Int32"))
        p_oDTMstr.Columns.Add("nPurPrice", Type.GetType("System.Decimal"))
    End Sub

    Private Function saveDetail(lnRow As Integer) As Boolean
        Dim lsTmpSQL As String
        Dim lsMasSQL As String, lsLgrSQL As String
        Dim lnQtyOnHnd As Integer, lnRpoOnHnd As Integer
        Dim lnLedgerNo As Integer, lbNewInventory As Boolean
        Dim loDta As DataTable

        Try
            With p_oDTMstr
                lsTmpSQL = "SELECT" &
                               "  a.*" &
                               ", IFNull(b.sBranchCd, '') xBranchCd" &
                               ", b.nQtyOnHnd xQtyOnHnd" &
                               ", b.nRpoOnHnd xRpoOnHnd" &
                               ", IFNULL(b.nLedgerNo, '') xLedgerNo" &
                            " FROM MC_Inventory a" &
                                  " LEFT JOIN MC_Inventory b" &
                                     " ON a.sMCInvIDx = b.sMCInvIDx" &
                                        " AND b.sBranchCd = " & strParm(p_sBranchCd) &
                            " WHERE a.sMCInvIDx = " & strParm(.Rows(lnRow).Item("sMCInvIDx")) &
                            " LIMIT 1"
                loDta = p_oApp.ExecuteQuery(lsTmpSQL)

                Select Case p_sSourceCd
                    Case pxeMCAcceptDelivery, pxeMCAcceptBackLoad, pxeMCImpound, pxeMCWarrantyPullOut
                        ' these are the only transaction that accepts existing motorcycle
                        '  so verify if an inventory exist for this branch
                        lbNewInventory = loDta(0).Item("xBranchCd") = ""
                End Select

                If lbNewInventory Then
                    lnQtyOnHnd = 0
                    lnRpoOnHnd = 0
                    lnLedgerNo = 1

                    lsMasSQL = "INSERT INTO MC_Inventory SET" & _
                                   "  sMCInvIDx = " & strParm(.Rows(lnRow).Item("sMCInvIDx")) & _
                                   ", sBranchCd = " & strParm(p_sBranchCd) & _
                                   ", sModelIDx = " & strParm(loDta(0).Item("sModelIDx")) & _
                                   ", sColorIDx = " & strParm(loDta(0).Item("sColorIDx")) & _
                                   ", dBegInvxx = " & dateParm(p_dTransact) & _
                                   ", nBegQtyxx = " & 0 & _
                                   ", nRpoQtyxx = " & 0 & _
                                   ", nQtyOnHnd = " & .Rows(lnRow).Item("nQtyInxxx") & _
                                   ", nRpoOnHnd = " & .Rows(lnRow).Item("nRpoQtyIn") & _
                                   ", nReOrderx = " & IFNull(loDta(0).Item("nReOrderx"), 0) & _
                                   ", nMinLevel = " & IFNull(loDta(0).Item("nMinLevel"), 0) & _
                                   ", nMaxLevel = " & IFNull(loDta(0).Item("nMaxLevel"), 0) & _
                                   ", nPurPrice = " & IFNull(loDta(0).Item("nPurPrice"), 0) & _
                                   ", nDealPrce = " & IIf(IsDBNull(loDta(0).Item("nDealPrce")), 0, loDta(0).Item("nDealPrce")) & _
                                   ", nSelPrice = " & IFNull(loDta(0).Item("nSelPrice"), 0) & _
                                   ", nSRPricex = " & IFNull(loDta(0).Item("nSRPricex"), 0) & _
                                   ", nLedgerNo = " & strParm(Format(lnLedgerNo, pxeLgrNoPict)) & _
                                   ", cRecdStat = " & strParm(GRider.xeLogical_YES) & _
                                   ", sModified = " & strParm(p_oApp.UserID) & _
                                   ", dModified = " & dateParm(p_oApp.getSysDate)
                Else
                    lnQtyOnHnd = IFNull(loDta(0).Item("xQtyOnHnd"), 0)
                    lnRpoOnHnd = IFNull(loDta(0).Item("xRpoOnHnd"), 0)
                    lnLedgerNo = IFNull(IIf(loDta(0).Item("xLedgerNo") = "", 0, loDta(0).Item("xLedgerNo")), 0) + 1

                    lsMasSQL = "UPDATE MC_Inventory SET" & _
                                   "  nQtyOnHnd = " & lnQtyOnHnd + .Rows(lnRow).Item("nQtyInxxx") - .Rows(lnRow).Item("nQtyOutxx") & _
                                   ", nRpoOnHnd = " & lnRpoOnHnd + .Rows(lnRow).Item("nRpoQtyIn") - .Rows(lnRow).Item("nRpoQtyOt") & _
                                   ", nLedgerNo = " & strParm(Format(lnLedgerNo, pxeLgrNoPict)) & _
                                   ", dModified = " & dateParm(p_oApp.getSysDate) & _
                                " WHERE sMCInvIDx = " & strParm(.Rows(lnRow).Item("sMCInvIDx")) & _
                                   " AND sBranchCd = " & strParm(p_sBranchCd)
                End If

                lsLgrSQL = "INSERT INTO MC_Inventory_Ledger SET" & _
                               "  sMCInvIDx = " & strParm(.Rows(lnRow).Item("sMCInvIDx")) & _
                               ", nLedgerNo = " & strParm(Format(lnLedgerNo, pxeLgrNoPict)) & _
                               ", sBranchCd = " & strParm(p_sBranchCd) & _
                               ", dTransact = " & dateParm(p_dTransact) & _
                               ", sSourceCd = " & strParm(p_sSourceCd) & _
                               ", sSourceNo = " & strParm(p_sSourceNo) & _
                               ", nQtyInxxx = " & .Rows(lnRow).Item("nQtyInxxx") & _
                               ", nQtyOutxx = " & .Rows(lnRow).Item("nQtyOutxx") & _
                               ", nRpoQtyIn = " & .Rows(lnRow).Item("nRpoQtyIn") & _
                               ", nRpoQtyOt = " & .Rows(lnRow).Item("nRpoQtyOt") & _
                               ", nQtyOnHnd = " & lnQtyOnHnd + .Rows(lnRow).Item("nQtyInxxx") - .Rows(lnRow).Item("nQtyOutxx") & _
                               ", nRpoOnHnd = " & lnRpoOnHnd + .Rows(lnRow).Item("nRpoQtyIn") - .Rows(lnRow).Item("nRpoQtyOt") & _
                               ", dModified = " & dateParm(p_oApp.getSysDate)
                Debug.Print(lsMasSQL)
                lnRow = p_oApp.Execute(lsMasSQL, "MC_Inventory", p_sBranchCd)
                If lnRow <= 0 Then
                    MsgBox(lsMasSQL & vbCrLf & _
                             "Unable to Update Motorcycle Inventory!", vbCritical, "Warning")
                    Return False
                End If

                lnRow = p_oApp.Execute(lsLgrSQL, "MC_Inventory_Ledger", p_sBranchCd)
                If lnRow <= 0 Then
                    MsgBox(lsLgrSQL & vbCrLf & _
                             "Unable to Update Motorcycle Inventory Transaction!", vbCritical, "Warning")
                    Return False
                End If
            End With
        Catch ex As Exception
            Throw ex
            Return False
        End Try

        saveDetail = True
    End Function

    Public Sub New(ByVal foRider As GRider)
        p_oApp = foRider
        p_sBranchCd = p_oApp.BranchCode
    End Sub
End Class

