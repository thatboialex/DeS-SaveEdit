Imports System.IO

Public Class DeS

    Public Shared bigendian As Boolean = False  ' PS5 is little-endian
    Public Shared bytes As Byte()
    Public Shared folder As String
    Public Shared filename As String
    Public Shared modified As Boolean = False

    Public Shared cllHairstyles As UInteger()
    Public Shared cllRings As UInteger()
    Public Shared cllWeapons As UInteger()
    Public Shared cllArmor As UInteger()
    Public Shared cllStartClass As UInteger()
    Public Shared cllItems As UInteger()
    Public Shared cllSpells As UInteger()
    Public Shared cllSpellStatus As String()

    Public Class WarpAreas
        Public Property name As String
        Public Property world As Integer
        Public Property block As Integer
        Public Property x As Single
        Public Property y As Single
        Public Property z As Single
        Public Property rot As Single
    End Class

    Public Shared cllWarps As New List(Of WarpAreas)
    Public Shared inventory(,) As UInteger

    ' =====================================================================
    ' FILE I/O - plain unencrypted read/write for PS5
    ' =====================================================================
    Shared Sub BytesToFile(name As String, b As Byte())
        IO.File.WriteAllBytes(folder & name, b)
    End Sub

    Shared Function FileToBytes(name As String) As Byte()
        Return IO.File.ReadAllBytes(folder & name)
    End Function

    ' =====================================================================
    ' READ HELPERS - little-endian (no byte reversal)
    ' =====================================================================
    Private Function RInt16(start As Integer) As Int16
        Dim ba(1) As Byte
        Array.Copy(bytes, start, ba, 0, 2)
        Return BitConverter.ToInt16(ba, 0)
    End Function

    Private Function RInt32(start As Integer) As Int32
        Dim ba(3) As Byte
        Array.Copy(bytes, start, ba, 0, 4)
        Return BitConverter.ToInt32(ba, 0)
    End Function

    Private Function RUInt16(start As Integer) As UInt16
        Dim ba(1) As Byte
        Array.Copy(bytes, start, ba, 0, 2)
        Return BitConverter.ToUInt16(ba, 0)
    End Function

    Private Function RUInt32(start As Integer) As UInteger
        Dim ba(3) As Byte
        Array.Copy(bytes, start, ba, 0, 4)
        Return BitConverter.ToUInt32(ba, 0)
    End Function

    Private Function RSingle(start As Integer) As Single
        Dim ba(3) As Byte
        Array.Copy(bytes, start, ba, 0, 4)
        Return BitConverter.ToSingle(ba, 0)
    End Function

    Private Function RUniStr(ByVal loc As UInteger, ByVal len As UInteger) As String
        Dim buildstr As String = ""
        For i As UInteger = 0 To len
            If bytes(loc + i) <> 0 Then
                buildstr = buildstr & Convert.ToChar(bytes(loc + i))
            Else
                If bytes(loc + i + 1) = 0 Then Exit For
            End If
        Next
        Return buildstr
    End Function

    ' =====================================================================
    ' WRITE HELPERS - little-endian (no byte reversal)
    ' =====================================================================
    Private Sub WBytes(ByVal loc As UInteger, ByVal byt As Byte())
        Array.Copy(byt, 0, bytes, loc, byt.Length)
    End Sub

    Private Sub WInt8(ByVal loc As UInteger, ByVal val As SByte)
        bytes(loc) = val
    End Sub

    Private Sub WInt16(ByVal loc As UInteger, ByVal val As Int16)
        WBytes(loc, BitConverter.GetBytes(val))
    End Sub

    Private Sub WInt32(ByVal loc As UInteger, ByVal val As Int32)
        WBytes(loc, BitConverter.GetBytes(val))
    End Sub

    Private Sub WUInt8(ByVal loc As UInteger, ByVal val As Byte)
        bytes(loc) = val
    End Sub

    Private Sub WUInt16(ByVal loc As UInteger, ByVal val As UInt16)
        WBytes(loc, BitConverter.GetBytes(val))
    End Sub

    Private Sub WUInt32(ByVal loc As UInteger, ByVal val As UInt32)
        WBytes(loc, BitConverter.GetBytes(val))
    End Sub

    Private Sub WSingle(ByVal loc As UInteger, ByVal val As Single)
        WBytes(loc, BitConverter.GetBytes(val))
    End Sub

    Private Function OneByteAnd(ByVal loc As UInteger, ByVal cmp As UInteger) As Boolean
        Return (bytes(loc) And cmp) <> 0
    End Function

    ' =====================================================================
    ' OPEN - loads quicksave## file directly from PS5 save folder
    ' =====================================================================
    Private Sub btnDeSOpen_Click(sender As System.Object, e As System.EventArgs) Handles btnDeSOpen.Click
        If Not txtDeSFolder.Text.EndsWith("") Then txtDeSFolder.Text = txtDeSFolder.Text & ""

        Try
            ' PS5 save slots: quicksave01 through quicksave10
            Dim slotNum As Integer = CInt(nmbSaveNum.Value)
            Dim slotName As String = "quicksave" & slotNum.ToString("D2")

            If Not IO.File.Exists(txtDeSFolder.Text & slotName) Then
                MsgBox("PS5 save file not found: " & slotName & vbCrLf & "Point the folder to the directory containing your quicksave## files.")
                Return
            End If

            txtF1.Text = slotName
            filename = slotName
            bytes = IO.File.ReadAllBytes(folder & filename)

            ' NOTE: All offsets below are from the PS3 version and marked
            ' [PS5-VERIFY] - confirm against your actual PS5 save in a hex editor
            ' before trusting the values.

            ' [PS5-VERIFY] World / position
            txtWorld.Text = Convert.ToUInt16(bytes(&H4))
            txtBlock.Text = Convert.ToUInt16(bytes(&H5))
            Dim posOffset As UInteger = RInt16(&H21AE1)
            txtXpos.Text = RSingle(&H21AE3 + posOffset)
            txtYpos.Text = RSingle(&H21AE3 + posOffset + 4)
            txtZpos.Text = RSingle(&H21AE3 + posOffset + 8)
            txtRot.Text = RSingle(&H21AE3 + posOffset + &H14)

            ' [PS5-VERIFY] Stats
            txtCurrHP.Text = RInt32(&H50)
            txtMaxHP.Text = RInt32(&H58)
            txtCurrMP.Text = RInt32(&H5C)
            txtMaxMP.Text = RInt32(&H64)
            txtCurrStam.Text = RInt32(&H6C)
            txtMaxStam.Text = RInt32(&H74)
            txtVit.Text = RInt32(&H80)
            txtInt.Text = RInt32(&H88)
            txtEnd.Text = RInt32(&H90)
            txtStr.Text = RInt32(&H98)
            txtDex.Text = RInt32(&HA0)
            txtMagic.Text = RInt32(&HA8)
            txtFaith.Text = RInt32(&HB0)
            txtLuck.Text = RInt32(&HB8)
            txtSouls.Text = RInt32(&HBC)
            txtSoulMem.Text = RInt32(&HC8)
            txtLevelsPurchased.Text = RInt32(&HCC)
            txtPhantomType.Text = bytes(&HD3)
            txtName.Text = RUniStr(&HD4, &H21)
            cmbGender.SelectedIndex = bytes(&HF6)
            cmbStartClass.SelectedIndex = bytes(&HFB)

            ' [PS5-VERIFY] Equipment slots
            cmbLeftHand1.SelectedIndex = Array.IndexOf(cllWeapons, RUInt32(&H28C))
            cmbRightHand1.SelectedIndex = Array.IndexOf(cllWeapons, RUInt32(&H290))
            cmbLeftHand2.SelectedIndex = Array.IndexOf(cllWeapons, RUInt32(&H294))
            cmbRightHand2.SelectedIndex = Array.IndexOf(cllWeapons, RUInt32(&H298))
            cmbArrows.SelectedIndex = Array.IndexOf(cllWeapons, RUInt32(&H29C))
            cmbBolts.SelectedIndex = Array.IndexOf(cllWeapons, RUInt32(&H2A0))
            cmbHelmet.SelectedIndex = Array.IndexOf(cllArmor, RUInt32(&H2A4))
            cmbChest.SelectedIndex = Array.IndexOf(cllArmor, RUInt32(&H2A8))
            cmbGauntlets.SelectedIndex = Array.IndexOf(cllArmor, RUInt32(&H2AC))
            cmbLeggings.SelectedIndex = Array.IndexOf(cllArmor, RUInt32(&H2B0))
            cmbHairstyle.SelectedIndex = Array.IndexOf(cllHairstyles, RUInt32(&H2B4))
            cmbRing1.SelectedIndex = Array.IndexOf(cllRings, RUInt32(&H2B8))
            cmbRing2.SelectedIndex = Array.IndexOf(cllRings, RUInt32(&H2BC))
            cmbQuickSlot1.SelectedIndex = Array.IndexOf(cllItems, RUInt32(&H2C0))
            cmbQuickSlot2.SelectedIndex = Array.IndexOf(cllItems, RUInt32(&H2C4))
            cmbQuickSlot3.SelectedIndex = Array.IndexOf(cllItems, RUInt32(&H2C8))
            cmbQuickSlot4.SelectedIndex = Array.IndexOf(cllItems, RUInt32(&H2CC))
            cmbQuickSlot5.SelectedIndex = Array.IndexOf(cllItems, RUInt32(&H2D0))

            ' [PS5-VERIFY] Inventory
            Dim invCount As UInteger = RUInt32(&H2D4)
            ReDim inventory(invCount, 8)
            dgvWeapons.Rows.Clear()
            dgvArmor.Rows.Clear()
            dgvRings.Rows.Clear()
            dgvGoods.Rows.Clear()
            dgvSpells.Rows.Clear()

            Dim i As UInteger
            Dim ItemType As Integer
            Dim ItemID As UInteger
            Dim ItemCount As UInteger
            Dim Idx1 As UInt16
            Dim Misc1 As UInteger
            Dim Idx2 As UInt16
            Dim Misc2 As UInteger
            Dim Durability As UInteger
            Dim offset As Integer = -&H20

            For i = 0 To invCount - 1
                ItemType = -1
                While ItemType = -1
                    offset += &H20
                    ItemType = RInt32(&H2DC + offset)
                End While
                ItemID = RUInt32(&H2E0 + offset)
                ItemCount = RUInt32(&H2E4 + offset)
                Idx1 = RUInt32(&H2E8 + offset)
                Misc1 = RUInt16(&H2EC + offset)
                Idx2 = RUInt16(&H2EE + offset)
                Misc2 = RUInt32(&H2F0 + offset)
                Durability = RUInt32(&H10364 + Idx1 * 8)

                Select Case ItemType
                    Case 0
                        dgvWeapons.Rows.Add(cmbLeftHand1.Items(Array.IndexOf(cllWeapons, ItemID)), ItemCount, Idx1, Misc1, Idx2, Misc2, Durability)
                    Case &H10000000
                        dgvArmor.Rows.Add(cmbChest.Items(Array.IndexOf(cllArmor, ItemID)), ItemCount, Idx1, Misc1, Idx2, Misc2, Durability)
                    Case &H20000000
                        dgvRings.Rows.Add(cmbRing1.Items(Array.IndexOf(cllRings, ItemID)), ItemCount, Idx1, Misc1, Idx2, Misc2, Durability)
                    Case &H40000000
                        dgvGoods.Rows.Add(cmbQuickSlot1.Items(Array.IndexOf(cllItems, ItemID)), ItemCount, Idx1, Misc1, Idx2, Misc2, Durability)
                End Select
            Next

            ' [PS5-VERIFY] Spells
            txtSpellSlots.Text = RUInt32(&H102E0)
            txtMiracleSlots.Text = RUInt32(&H1030C)
            txtHairR.Text = Math.Round(Val(RSingle(&H14368)), 3)
            txtHairG.Text = Math.Round(Val(RSingle(&H1436C)), 3)
            txtHairB.Text = Math.Round(Val(RSingle(&H14370)), 3)

            Dim spellCount As UInteger = RUInt32(&H143E8)
            For i = 0 To spellCount - 1
                Dim spellStatus As UInteger = RUInt32(&H143EC + i * &H10)
                Dim spellID As UInteger = RUInt32(&H143F0 + i * &H10)
                Misc1 = RUInt32(&H143F4 + i * &H10)
                Misc2 = RUInt32(&H143F8 + i * &H10)
                dgvSpells.Rows.Add(cmbSpellSlot1.Items(Array.IndexOf(cllSpells, spellID)), cllSpellStatus(spellStatus), Misc1, Misc2)
            Next

            ' [PS5-VERIFY] Tendency + NPC flags
            txtCharTendency.Text = RSingle(&H1EBF0)
            txtNexusTendency.Text = RSingle(&H1EBF8)
            txtW1Tendency.Text = RSingle(&H1EC00)
            txtW2Tendency.Text = RSingle(&H1EC20)
            txtW3Tendency.Text = RSingle(&H1EC10)
            txtW4Tendency.Text = RSingle(&H1EC08)
            txtW5Tendency.Text = RSingle(&H1EC18)
            txtClearCount.Text = bytes(&H1EC58)
            chkArchSealed.Checked = Not OneByteAnd(&H1F965, &H40)
            chkEdWorking.Checked = OneByteAnd(&H1FD55, &H4)
            chkEdHostile.Checked = OneByteAnd(&H1FD55, &H8)
            chkEdDead.Checked = OneByteAnd(&H1FD55, &H10)
            chkThomasFriendly.Checked = OneByteAnd(&H1FD75, &H40)
            chkThomasHostile.Checked = OneByteAnd(&H1FD75, &H80)
            chkThomasDead.Checked = OneByteAnd(&H1FD76, &H1)
            chkBoldwinFriendly.Checked = OneByteAnd(&H1FD81, &H1)
            chkBoldwinHostile.Checked = OneByteAnd(&H1FD81, &H2)
            chkBoldwinDead.Checked = OneByteAnd(&H1FD81, &H4)

            cmbLocations.SelectedIndex = 0

        Catch ex As Exception
            MsgBox("Failed to open save: " & ex.Message)
        End Try
    End Sub

    ' =====================================================================
    ' SAVE - writes directly back to the quicksave## file
    ' =====================================================================
    Private Sub btnDeSSave_Click(sender As System.Object, e As System.EventArgs) Handles btnDeSSave.Click
        Try
            filename = txtF1.Text
            bytes = FileToBytes(filename)

            ' [PS5-VERIFY] World / position
            bytes(&H4) = Val(txtWorld.Text)
            bytes(&H5) = Val(txtBlock.Text)
            bytes(&H6) = 0
            bytes(&H7) = 0
            Dim posOffset As UInteger = RInt16(&H21AE1)
            WSingle(&H21AE3 + posOffset, Val(txtXpos.Text))
            WSingle(&H21AE3 + posOffset + 4, Val(txtYpos.Text))
            WSingle(&H21AE3 + posOffset + 8, Val(txtZpos.Text))
            WSingle(&H21AE3 + posOffset + &H14, Val(txtRot.Text))

            ' [PS5-VERIFY] Stats
            WInt32(&H50, Val(txtCurrHP.Text))
            WInt32(&H58, Val(txtMaxHP.Text))
            WInt32(&H5C, Val(txtCurrMP.Text))
            WInt32(&H64, Val(txtMaxMP.Text))
            WInt32(&H6C, Val(txtCurrStam.Text))
            WInt32(&H74, Val(txtMaxStam.Text))
            WInt32(&H7C, Val(txtVit.Text))
            WInt32(&H80, Val(txtVit.Text))
            WInt32(&H84, Val(txtInt.Text))
            WInt32(&H88, Val(txtInt.Text))
            WInt32(&H8C, Val(txtEnd.Text))
            WInt32(&H90, Val(txtEnd.Text))
            WInt32(&H94, Val(txtStr.Text))
            WInt32(&H98, Val(txtStr.Text))
            WInt32(&H9C, Val(txtDex.Text))
            WInt32(&HA0, Val(txtDex.Text))
            WInt32(&HA4, Val(txtMagic.Text))
            WInt32(&HA8, Val(txtMagic.Text))
            WInt32(&HAC, Val(txtFaith.Text))
            WInt32(&HB0, Val(txtFaith.Text))
            WInt32(&HB4, Val(txtLuck.Text))
            WInt32(&HB8, Val(txtLuck.Text))
            WInt32(&HBC, Val(txtSouls.Text))
            WInt32(&HC8, Val(txtSoulMem.Text))
            WInt32(&HCC, Val(txtLevelsPurchased.Text))
            bytes(&HD3) = Val(txtPhantomType.Text)

            ' Name
            For i As Integer = 0 To &H10
                If i < txtName.Text.Length Then
                    bytes(&HD5 + i * 2) = Microsoft.VisualBasic.Asc(txtName.Text(i))
                Else
                    bytes(&HD5 + i * 2) = 0
                End If
                bytes(&HD5 + i * 2 + 1) = 0
            Next

            bytes(&HF6) = cmbGender.SelectedIndex
            bytes(&HFB) = cmbStartClass.SelectedIndex

            ' [PS5-VERIFY] Equipment
            WUInt32(&H28C, Val(cllWeapons(cmbLeftHand1.SelectedIndex)))
            WUInt32(&H290, Val(cllWeapons(cmbRightHand1.SelectedIndex)))
            WUInt32(&H294, Val(cllWeapons(cmbLeftHand2.SelectedIndex)))
            WUInt32(&H298, Val(cllWeapons(cmbRightHand2.SelectedIndex)))
            WUInt32(&H29C, Val(cllWeapons(cmbArrows.SelectedIndex)))
            WUInt32(&H2A0, Val(cllWeapons(cmbBolts.SelectedIndex)))
            WUInt32(&H2A4, Val(cllArmor(cmbHelmet.SelectedIndex)))
            WUInt32(&H2A8, Val(cllArmor(cmbChest.SelectedIndex)))
            WUInt32(&H2AC, Val(cllArmor(cmbGauntlets.SelectedIndex)))
            WUInt32(&H2B0, Val(cllArmor(cmbLeggings.SelectedIndex)))
            WUInt32(&H2B4, Val(cllHairstyles(cmbHairstyle.SelectedIndex)))
            WUInt32(&H2B8, Val(cllRings(cmbRing1.SelectedIndex)))
            WUInt32(&H2BC, Val(cllRings(cmbRing2.SelectedIndex)))

            ' [PS5-VERIFY] Inventory write-back
            Dim invCount As Integer = Val(dgvWeapons.Rows.Count) + Val(dgvArmor.Rows.Count) + Val(dgvRings.Rows.Count) + Val(dgvGoods.Rows.Count) - 4
            WUInt32(&H2D4, invCount)
            WUInt32(&H10360, invCount)

            For i As Integer = 0 To &H800 - 1
                WInt32(&H2DC + i * &H20, -1)
                WInt32(&H2E0 + i * &H20, -1)
                WInt32(&H2E4 + i * &H20, 0)
                WInt32(&H2E8 + i * &H20, -1)
                WInt16(&H2EC + i * &H20, -1)
                WInt16(&H2EE + i * &H20, -1)
                WInt32(&H2F0 + i * &H20, 0)
                WInt32(&H2F4 + i * &H20, 0)
                WInt32(&H2F8 + i * &H20, 0)
            Next

            Dim invslot As Integer = 0

            If dgvWeapons.Rows.Count > 0 Then
                For i As Integer = 0 To dgvWeapons.Rows.Count - 2
                    WUInt32(&H2DC + invslot * &H20, 0)
                    WUInt32(&H2E0 + invslot * &H20, Val(cllWeapons(cmbLeftHand1.FindStringExact(dgvWeapons.Rows(i).Cells(0).FormattedValue))))
                    WUInt32(&H2E4 + invslot * &H20, Val(dgvWeapons.Rows(i).Cells(1).FormattedValue))
                    WUInt32(&H2E8 + invslot * &H20, Val(dgvWeapons.Rows(i).Cells(2).FormattedValue))
                    WUInt16(&H2EC + invslot * &H20, Val(dgvWeapons.Rows(i).Cells(3).FormattedValue))
                    WUInt16(&H2EE + invslot * &H20, Val(dgvWeapons.Rows(i).Cells(4).FormattedValue))
                    WUInt32(&H2F0 + invslot * &H20, Val(dgvWeapons.Rows(i).Cells(5).FormattedValue))
                    WUInt32(&H10364 + Val(dgvWeapons.Rows(i).Cells(2).FormattedValue) * 8, Val(dgvWeapons.Rows(i).Cells(6).FormattedValue))
                    invslot += 1
                Next
            End If

            If dgvArmor.Rows.Count > 0 Then
                For i As Integer = 0 To dgvArmor.Rows.Count - 2
                    WUInt32(&H2DC + invslot * &H20, &H10000000)
                    WUInt32(&H2E0 + invslot * &H20, Val(cllArmor(cmbChest.FindStringExact(dgvArmor.Rows(i).Cells(0).FormattedValue))))
                    WUInt32(&H2E4 + invslot * &H20, Val(dgvArmor.Rows(i).Cells(1).FormattedValue))
                    WUInt32(&H2E8 + invslot * &H20, Val(dgvArmor.Rows(i).Cells(2).FormattedValue))
                    WUInt16(&H2EC + invslot * &H20, Val(dgvArmor.Rows(i).Cells(3).FormattedValue))
                    WUInt16(&H2EE + invslot * &H20, Val(dgvArmor.Rows(i).Cells(4).FormattedValue))
                    WUInt32(&H2F0 + invslot * &H20, Val(dgvArmor.Rows(i).Cells(5).FormattedValue))
                    WUInt32(&H10364 + Val(dgvArmor.Rows(i).Cells(2).FormattedValue) * 8, Val(dgvArmor.Rows(i).Cells(6).FormattedValue))
                    invslot += 1
                Next
            End If

            If dgvRings.Rows.Count > 0 Then
                For i As Integer = 0 To dgvRings.Rows.Count - 2
                    WUInt32(&H2DC + invslot * &H20, &H20000000)
                    WUInt32(&H2E0 + invslot * &H20, Val(cllRings(cmbRing1.FindStringExact(dgvRings.Rows(i).Cells(0).FormattedValue))))
                    WUInt32(&H2E4 + invslot * &H20, Val(dgvRings.Rows(i).Cells(1).FormattedValue))
                    WUInt32(&H2E8 + invslot * &H20, Val(dgvRings.Rows(i).Cells(2).FormattedValue))
                    WUInt16(&H2EC + invslot * &H20, Val(dgvRings.Rows(i).Cells(3).FormattedValue))
                    WUInt16(&H2EE + invslot * &H20, Val(dgvRings.Rows(i).Cells(4).FormattedValue))
                    WUInt32(&H2F0 + invslot * &H20, Val(dgvRings.Rows(i).Cells(5).FormattedValue))
                    invslot += 1
                Next
            End If

            If dgvGoods.Rows.Count > 0 Then
                For i As Integer = 0 To dgvGoods.Rows.Count - 2
                    WUInt32(&H2DC + invslot * &H20, &H40000000)
                    WUInt32(&H2E0 + invslot * &H20, Val(cllItems(cmbQuickSlot1.FindStringExact(dgvGoods.Rows(i).Cells(0).FormattedValue))))
                    WUInt32(&H2E4 + invslot * &H20, Val(dgvGoods.Rows(i).Cells(1).FormattedValue))
                    WUInt32(&H2E8 + invslot * &H20, Val(dgvGoods.Rows(i).Cells(2).FormattedValue))
                    WUInt16(&H2EC + invslot * &H20, Val(dgvGoods.Rows(i).Cells(3).FormattedValue))
                    WUInt16(&H2EE + invslot * &H20, Val(dgvGoods.Rows(i).Cells(4).FormattedValue))
                    WUInt32(&H2F0 + invslot * &H20, Val(dgvGoods.Rows(i).Cells(5).FormattedValue))
                    invslot += 1
                Next
            End If

            ' [PS5-VERIFY] Spells
            WUInt32(&H102E0, Val(Val(txtSpellSlots.Text)))
            WUInt32(&H1030C, Val(Val(txtMiracleSlots.Text)))
            WUInt32(&H143E8, Val(dgvSpells.Rows.Count) - 1)
            invslot = 0
            If dgvSpells.Rows.Count > 0 Then
                For i As Integer = 0 To dgvSpells.Rows.Count - 2
                    WUInt32(&H143EC + invslot * &H10, Val(Array.IndexOf(cllSpellStatus, dgvSpells.Rows(i).Cells(1).FormattedValue)))
                    WUInt32(&H143F0 + invslot * &H10, Val(cllSpells(cmbSpellSlot1.FindStringExact(dgvSpells.Rows(i).Cells(0).FormattedValue))))
                    WUInt32(&H143F4 + invslot * &H10, Val(dgvSpells.Rows(i).Cells(2).FormattedValue))
                    WUInt32(&H143F8 + invslot * &H10, Val(dgvSpells.Rows(i).Cells(3).FormattedValue))
                    invslot += 1
                Next
            End If

            ' [PS5-VERIFY] Appearance
            WSingle(&H14368, Val(txtHairR.Text))
            WSingle(&H1436C, Val(txtHairG.Text))
            WSingle(&H14370, Val(txtHairB.Text))

            ' [PS5-VERIFY] Tendency
            WSingle(&H1EBF0, Val(txtCharTendency.Text))
            WSingle(&H1EBF8, Val(txtNexusTendency.Text))
            WSingle(&H1EBFC, Val(txtNexusTendency.Text))
            WSingle(&H1EC00, Val(txtW1Tendency.Text))
            WSingle(&H1EC04, Val(txtW1Tendency.Text))
            WSingle(&H1EC08, Val(txtW4Tendency.Text))
            WSingle(&H1EC0C, Val(txtW4Tendency.Text))
            WSingle(&H1EC10, Val(txtW3Tendency.Text))
            WSingle(&H1EC14, Val(txtW3Tendency.Text))
            WSingle(&H1EC18, Val(txtW5Tendency.Text))
            WSingle(&H1EC1C, Val(txtW5Tendency.Text))
            WSingle(&H1EC20, Val(txtW2Tendency.Text))
            WSingle(&H1EC24, Val(txtW2Tendency.Text))
            bytes(&H1EC58) = Val(txtClearCount.Text)

            ' [PS5-VERIFY] NPC flags
            bytes(&H1F965) = (bytes(&H1F965) And &HBF) Or (IIf(Not chkArchSealed.Checked, &H40, 0))
            bytes(&H1FD55) = (bytes(&H1FD55) And &HFB) Or (IIf(chkEdWorking.Checked, &H4, 0))
            bytes(&H1FD55) = (bytes(&H1FD55) And &HF7) Or (IIf(chkEdHostile.Checked, &H8, 0))
            bytes(&H1FD55) = (bytes(&H1FD55) And &HEF) Or (IIf(chkEdDead.Checked, &H10, 0))
            bytes(&H1FD75) = (bytes(&H1FD75) And &HBF) Or (IIf(chkThomasFriendly.Checked, &H40, 0))
            bytes(&H1FD75) = (bytes(&H1FD75) And &H7F) Or (IIf(chkThomasHostile.Checked, &H80, 0))
            bytes(&H1FD76) = (bytes(&H1FD76) And &HFE) Or (IIf(chkThomasDead.Checked, &H1, 0))
            bytes(&H1FD81) = (bytes(&H1FD81) And &HFE) Or (IIf(chkBoldwinFriendly.Checked, &H1, 0))
            bytes(&H1FD81) = (bytes(&H1FD81) And &HFD) Or (IIf(chkBoldwinHostile.Checked, &H2, 0))
            bytes(&H1FD81) = (bytes(&H1FD81) And &HFB) Or (IIf(chkBoldwinDead.Checked, &H4, 0))

            BytesToFile(filename, bytes)
            MsgBox("Save Completed")

        Catch ex As Exception
            MsgBox("Save failed: " & ex.Message)
        End Try
    End Sub

    ' =====================================================================
    ' BROWSE - pick the folder containing quicksave## files
    ' =====================================================================
    Private Sub btnDeSBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnDeSBrowse.Click
        Dim openDlg As New OpenFileDialog
        openDlg.Filter = "PS5 DeS Save|quicksave*|All Files|*.*"
        openDlg.Title = "Select any quicksave## file in your PS5 save folder..."
        If openDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtDeSFolder.Text = Microsoft.VisualBasic.Left(openDlg.FileName, InStrRev(openDlg.FileName, ""))
        End If
    End Sub

    Private Sub txtDeSFolder_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDeSFolder.TextChanged
        folder = UCase(txtDeSFolder.Text)
    End Sub

    Private Sub cmbLocations_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbLocations.SelectedIndexChanged
        Dim area As WarpAreas = cllWarps.Item(cmbLocations.SelectedIndex)
        txtWorld.Text = area.world
        txtBlock.Text = area.block
        txtXpos.Text = area.x
        txtYpos.Text = area.y
        txtZpos.Text = area.z
        txtRot.Text = area.rot
    End Sub

    Private Sub txtWorld_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtWorld.TextChanged
        Select Case txtWorld.Text
            Case "1" : lblWorldName.Text = "The Nexus"
            Case "2" : lblWorldName.Text = "Boletarian Palace"
            Case "3" : lblWorldName.Text = "Shrine of Storms"
            Case "4" : lblWorldName.Text = "Tower of Latria"
            Case "5" : lblWorldName.Text = "Valley of Defilement"
            Case "6" : lblWorldName.Text = "Stonefang Tunnel"
            Case "7" : lblWorldName.Text = "Northern Limit"
            Case "8" : lblWorldName.Text = "Tutorial"
            Case Else : lblWorldName.Text = "unknown"
        End Select
    End Sub

    Private Sub btnSendMoney_Click(sender As Object, e As EventArgs) Handles btnSendMoney.Click
        System.Diagnostics.Process.Start("https://www.paypal.me/wulf2k")
    End Sub

    ' =====================================================================
    ' InitArrays - PS5 IDs from debugmenu.txt
    ' =====================================================================
    Sub InitArrays()
        cllWarps.Add(New WarpAreas With {.name = "Nexus", .world = 1, .block = 0, .x = 0, .y = 0, .z = 0, .rot = 0})
        cllWarps.Add(New WarpAreas With {.name = "Gates of Boletaria, Start", .world = 2, .block = 0, .x = 17.5, .y = 21.2, .z = -159.3, .rot = 3})
        cllWarps.Add(New WarpAreas With {.name = "Gates of Boletaria, Phalanx", .world = 2, .block = 0, .x = 18, .y = 30.3, .z = -45.7, .rot = 3})
        cllWarps.Add(New WarpAreas With {.name = "The Lords Path, Tower Knight", .world = 2, .block = 1, .x = 16, .y = 48.1, .z = 418, .rot = -3.1})
        cllWarps.Add(New WarpAreas With {.name = "Inner Ward, Penetrator", .world = 2, .block = 2, .x = -126.4, .y = 89.9, .z = 699, .rot = -2.73})
        cllWarps.Add(New WarpAreas With {.name = "The Kings Tower, King Allant", .world = 2, .block = 3, .x = -10.8, .y = 137.3, .z = 989.6, .rot = 3.13})
        cllWarps.Add(New WarpAreas With {.name = "Islands Edge, Adjudicator", .world = 3, .block = 1, .x = -5.2, .y = 100, .z = -421, .rot = 0.08})
        cllWarps.Add(New WarpAreas With {.name = "Altar of Storms, Storm King", .world = 3, .block = 3, .x = -4.8, .y = 6.7, .z = -450.9, .rot = 0})
        cllWarps.Add(New WarpAreas With {.name = "Prison of Hope, Fools Idol", .world = 4, .block = 0, .x = 183.8, .y = -9, .z = -91, .rot = -1.58})
        cllWarps.Add(New WarpAreas With {.name = "Upper Latria, Old Monk", .world = 4, .block = 2, .x = 569.7, .y = 158.8, .z = -160.2, .rot = 0.44})
        cllWarps.Add(New WarpAreas With {.name = "Depraved Chasm, Leechmonger", .world = 5, .block = 0, .x = 1.5, .y = 15.3, .z = 3.8, .rot = -0.26})
        cllWarps.Add(New WarpAreas With {.name = "Swamp of Sorrow, Dirty Colossus", .world = 5, .block = 1, .x = 99, .y = -21.9, .z = -395.5, .rot = 0.35})
        cllWarps.Add(New WarpAreas With {.name = "Rotting Haven, Maiden Astrea", .world = 5, .block = 2, .x = 92.6, .y = -38.7, .z = -476.4, .rot = -1.48})
        cllWarps.Add(New WarpAreas With {.name = "Smithing Grounds, Armor Spider", .world = 6, .block = 0, .x = 177.5, .y = 0, .z = 172.3, .rot = 0.14})
        cllWarps.Add(New WarpAreas With {.name = "The Tunnel City, Flamelurker", .world = 6, .block = 1, .x = 139.1, .y = -83.4, .z = -43.8, .rot = 0.07})
        cllWarps.Add(New WarpAreas With {.name = "Underground Temple, Dragon God", .world = 6, .block = 2, .x = 139.5, .y = -90.4, .z = -174.9, .rot = -0.09})

        cmbLocations.Items.Clear()
        For Each warp In cllWarps
            cmbLocations.Items.Add(warp.name)
        Next

        cllSpells = New UInteger() {
            1000, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009,
            1010, 1011, 1012, 1013, 1014, 1015, 1016, 1017, 1018, 1019,
            1020, 1021,
            2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009,
            2010, 2011,
            &HFFFFFFFF
        }

        cllRings = New UInteger() {
            0,
            100, 101, 102, 103, 104, 105, 106, 107, 108, 109,
            110, 111, 112, 113, 114, 115, 116, 117, 118, 119,
            120, 121, 123, 124, 125, 126,
            9960, 9961, 9962, 9963,
            &HFFFFFFFF
        }

        cllHairstyles = New UInteger() {
            &H7A120, &H7A184, &H7A1E8, &H7A24C, &H7A2B0, &H7A314, &H7A378,
            &H7A3DC, &H7A440, &H7A4A4, &H7A508, &H7A56C, &H7A5D0, &H7A634,
            &H7A698, &H7A6FC, &H7A760, &H7A7C4, &H7A828, &H7A88C, &H900B0,
            &HFFFFFFFF
        }

        cllArmor = New UInteger() {
            100000, 100100, 100200, 100300, 100400, 100500, 100600, 100700,
            100800, 100900, 101000, 101100, 101400, 101800, 101900, 102000,
            102100, 180000, 190400, 199905, 199906, 199907, 199909,
            200000, 200100, 200200, 200300, 200400, 200500, 200600, 200700,
            200800, 200900, 201000, 201100, 201200, 201400, 201800, 201900,
            202000, 202100, 299900, 299903, 299904, 299905, 299906, 299907,
            299908, 299909,
            300000, 300100, 300200, 300300, 300400, 300500, 300600, 300700,
            300800, 300900, 301000, 301100, 301200, 301400, 301700, 301800,
            301900, 302000, 302100, 399900, 399903, 399904, 399905, 399906,
            399907, 399908, 399909,
            400000, 400100, 400200, 400300, 400400, 400500, 400600, 400700,
            400800, 400900, 401000, 401100, 401200, 401400, 401800, 401900,
            402000, 402100, 499900, 499903, 499904, 499905, 499906, 499907,
            499908, 499909,
            &HFFFFFFFF
        }

        cllWeapons = New UInteger() {
            10000, 10100, 10200, 10400, 10500, 10600, 10700,
            20000, 20100, 20200, 20800, 21300, 21400, 21700, 21800,
            20300, 20400, 20500, 20900, 21000, 21100, 21500, 21600, 21900,
            20600, 20700,
            40000, 40100, 40200, 40300, 40400, 40500, 40600, 40700, 40900,
            50000, 50200, 50300, 50400, 50500,
            60000, 60100, 60200, 60300, 60400, 60500, 60600, 60700,
            70000, 70100, 70300, 70400,
            80000, 80100, 80200, 80300,
            90000, 90100, 90200, 90400, 90500,
            100000, 100100, 100200,
            130000, 130100, 130200, 130300, 130400, 130500,
            140000, 140100, 140200,
            150000, 150100, 150200, 150300, 150400, 150500, 150600, 150700,
            150800, 150900, 151000, 151100, 151200, 151300, 151400, 151500,
            160000, 160100, 160200, 160400, 160500, 160600, 160700, 160800,
            170000, 170100, 170200, 170300,
            999000, 999100, 999200,
            &HFFFFFFFF
        }

        cllItems = New UInteger() {
            1000, 1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1009,
            1010, 1011, 1012, 1013, 1014, 1015, 1016, 1017, 1020, 1021,
            1022, 1023, 1024,
            8, 9, 10, 12, 13, 14, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44,
            99, 1018, 1033, 9993, 9931,
            9995, 9996, 9997, 9999,
            7, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28,
            29, 30, 31, 32, 33, 34, 45,
            1025, 1026, 1027, 1028, 1029, 1030, 1031, 1032,
            2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2012, 2013,
            2014, 2015, 2016, 2017, 2021, 2022, 2023, 2024, 2025, 2026,
            2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035, 2036,
            2037, 2038, 2039, 2040, 2041, 2042, 2043, 2044, 2045, 2046,
            2047, 2048, 2049, 2050, 2054,
            9920, 9921, 9922, 9924, 9925, 9926, 9927, 9928, 9929, 9930,
            &HFFFFFFFF
        }

        cllSpellStatus = New String() {"Unavailable", "Unknown", "Known", "Memorized"}

        InitWeaps(cmbLeftHand1)
        InitWeaps(cmbLeftHand2)
        InitWeaps(cmbRightHand1)
        InitWeaps(cmbRightHand2)
        InitWeaps(cmbArrows)
        InitWeaps(cmbBolts)
        InitSpells(cmbSpellSlot1)

        dgvBuild(dgvWeapons, cmbLeftHand1, "Weapons")
        dgvBuild(dgvArmor, cmbChest, "Armor")
        dgvBuild(dgvRings, cmbRing1, "Rings")
        dgvBuild(dgvGoods, cmbQuickSlot1, "Goods")
        dgvSpellBuild(dgvSpells, cmbSpellSlot1, "Spells/Miracles")
    End Sub

    Sub InitWeaps(cmbbox As ComboBox)
        cmbbox.Items.Clear()
        cmbbox.Items.Add("Dagger")
        cmbbox.Items.Add("Parrying Dagger")
        cmbbox.Items.Add("Mail Breaker")
        cmbbox.Items.Add("Baby''s Nail")
        cmbbox.Items.Add("Geri''s Stiletto")
        cmbbox.Items.Add("Kris Blade")
        cmbbox.Items.Add("Secret Dagger")
        cmbbox.Items.Add("Short Sword")
        cmbbox.Items.Add("Broad Sword")
        cmbbox.Items.Add("Long Sword")
        cmbbox.Items.Add("Rune Sword")
        cmbbox.Items.Add("Knight Sword")
        cmbbox.Items.Add("Broken Sword")
        cmbbox.Items.Add("Blueblood Sword")
        cmbbox.Items.Add("Penetrating Sword")
        cmbbox.Items.Add("Flamberge")
        cmbbox.Items.Add("Bastard Sword")
        cmbbox.Items.Add("Claymore")
        cmbbox.Items.Add("Soulbrandt")
        cmbbox.Items.Add("Demonbrandt")
        cmbbox.Items.Add("Storm Ruler")
        cmbbox.Items.Add("Northern Regalia")
        cmbbox.Items.Add("Large Sword of Moonlight")
        cmbbox.Items.Add("Morion Blade")
        cmbbox.Items.Add("Great Sword")
        cmbbox.Items.Add("Dragon Bone Smasher")
        cmbbox.Items.Add("Scimitar")
        cmbbox.Items.Add("Kilij")
        cmbbox.Items.Add("Shotel")
        cmbbox.Items.Add("Falchion")
        cmbbox.Items.Add("Uchigatana")
        cmbbox.Items.Add("Hiltless")
        cmbbox.Items.Add("Blind")
        cmbbox.Items.Add("Magic Sword Makoto")
        cmbbox.Items.Add("Large Sword of Searching")
        cmbbox.Items.Add("Battle Axe")
        cmbbox.Items.Add("Great Axe")
        cmbbox.Items.Add("Crescent Axe")
        cmbbox.Items.Add("Guillotine Axe")
        cmbbox.Items.Add("Dozer Axe")
        cmbbox.Items.Add("Club")
        cmbbox.Items.Add("Mace")
        cmbbox.Items.Add("War Pick")
        cmbbox.Items.Add("Morning Star")
        cmbbox.Items.Add("Great Club")
        cmbbox.Items.Add("Bramd")
        cmbbox.Items.Add("Pickaxe")
        cmbbox.Items.Add("Meat Cleaver")
        cmbbox.Items.Add("Short Spear")
        cmbbox.Items.Add("Winged Spear")
        cmbbox.Items.Add("Istarelle")
        cmbbox.Items.Add("Scraping Spear")
        cmbbox.Items.Add("War Scythe")
        cmbbox.Items.Add("Mirdan Hammer")
        cmbbox.Items.Add("Halberd")
        cmbbox.Items.Add("Phosphorescent Pole")
        cmbbox.Items.Add("Wooden Catalyst")
        cmbbox.Items.Add("Silver Catalyst")
        cmbbox.Items.Add("Insanity Catalyst")
        cmbbox.Items.Add("Talisman of God")
        cmbbox.Items.Add("Talisman of Beasts")
        cmbbox.Items.Add("Iron Knuckles")
        cmbbox.Items.Add("Claws")
        cmbbox.Items.Add("Hands of God")
        cmbbox.Items.Add("Short Bow")
        cmbbox.Items.Add("Compound Short Bow")
        cmbbox.Items.Add("Long Bow")
        cmbbox.Items.Add("Compound Long Bow")
        cmbbox.Items.Add("White Bow")
        cmbbox.Items.Add("Lava Bow")
        cmbbox.Items.Add("Light Crossbow")
        cmbbox.Items.Add("Heavy Crossbow")
        cmbbox.Items.Add("Gargoyle Crossbow")
        cmbbox.Items.Add("Buckler")
        cmbbox.Items.Add("Wooden Shield")
        cmbbox.Items.Add("Kite Shield")
        cmbbox.Items.Add("Heater Shield")
        cmbbox.Items.Add("Adjudicator''s Shield")
        cmbbox.Items.Add("Spiked Shield")
        cmbbox.Items.Add("Tower Shield")
        cmbbox.Items.Add("Dark Silver Shield")
        cmbbox.Items.Add("Soldier''s Shield")
        cmbbox.Items.Add("Knight''s Shield")
        cmbbox.Items.Add("Slave''s Shield")
        cmbbox.Items.Add("Rune Shield")
        cmbbox.Items.Add("Large Brushwood Shield")
        cmbbox.Items.Add("Steel Shield")
        cmbbox.Items.Add("Purple Flame Shield")
        cmbbox.Items.Add("Leather Shield")
        cmbbox.Items.Add("Arrow")
        cmbbox.Items.Add("Heavy Arrow")
        cmbbox.Items.Add("Light Arrow")
        cmbbox.Items.Add("Fire Arrow")
        cmbbox.Items.Add("Rotten Arrow")
        cmbbox.Items.Add("Holy Arrow")
        cmbbox.Items.Add("White Arrow")
        cmbbox.Items.Add("Wooden Arrow")
        cmbbox.Items.Add("Bolt")
        cmbbox.Items.Add("Heavy Bolt")
        cmbbox.Items.Add("Black Bolt")
        cmbbox.Items.Add("Wooden Bolt")
        cmbbox.Items.Add("Reaper Scythe")
        cmbbox.Items.Add("Ritual Blade")
        cmbbox.Items.Add("Hoplite Shield")
        cmbbox.Items.Add("---No Weapon---")
    End Sub

    Sub InitSpells(cmbbox As ComboBox)
        cmbbox.Items.Clear()
        cmbbox.Items.Add("Soul Arrow")
        cmbbox.Items.Add("Flame Toss")
        cmbbox.Items.Add("Relief")
        cmbbox.Items.Add("Enchant Weapon")
        cmbbox.Items.Add("Curse Weapon")
        cmbbox.Items.Add("Soul Thirst")
        cmbbox.Items.Add("Poison Cloud")
        cmbbox.Items.Add("Demon''s Prank")
        cmbbox.Items.Add("Fireball")
        cmbbox.Items.Add("Ignite")
        cmbbox.Items.Add("Soul Ray")
        cmbbox.Items.Add("Homing Soul Arrow")
        cmbbox.Items.Add("Cloak")
        cmbbox.Items.Add("Protection")
        cmbbox.Items.Add("Light Weapon")
        cmbbox.Items.Add("Water Veil")
        cmbbox.Items.Add("Death Cloud")
        cmbbox.Items.Add("Fire Spray")
        cmbbox.Items.Add("Soulsucker")
        cmbbox.Items.Add("Acid Cloud")
        cmbbox.Items.Add("Warding")
        cmbbox.Items.Add("Firestorm")
        cmbbox.Items.Add("God''s Wrath")
        cmbbox.Items.Add("Anti-Magic Field")
        cmbbox.Items.Add("Recovery")
        cmbbox.Items.Add("Second Chance")
        cmbbox.Items.Add("Regeneration")
        cmbbox.Items.Add("Resurrection")
        cmbbox.Items.Add("Cure")
        cmbbox.Items.Add("Hidden Soul")
        cmbbox.Items.Add("Evacuate")
        cmbbox.Items.Add("Banish")
        cmbbox.Items.Add("Heal")
        cmbbox.Items.Add("Antidote")
    End Sub

    Private Sub dgvDefaultValuesNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) _
        Handles dgvWeapons.DefaultValuesNeeded, dgvArmor.DefaultValuesNeeded, dgvRings.DefaultValuesNeeded, dgvGoods.DefaultValuesNeeded
        With e.Row
            .Cells("Count").Value = 1
            .Cells("Misc1").Value = 0
            .Cells("Idx1").Value = 999
            .Cells("Misc2").Value = &H1000000
            .Cells("Idx2").Value = 999
            .Cells("Durability").Value = 999
        End With
    End Sub

    Private Sub dgvSpellsDefaultValuesNeeded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowEventArgs) _
        Handles dgvSpells.DefaultValuesNeeded
        With e.Row
            .Cells("Status").Value = "Unavailable"
            .Cells("Misc1").Value = 0
            .Cells("Misc2").Value = 0
        End With
    End Sub

    Private Sub txtDrop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtDeSFolder.DragDrop
        Dim f As String() = e.Data.GetData(DataFormats.FileDrop)
        sender.Text = Microsoft.VisualBasic.Left(f(0), InStrRev(f(0), ""))
    End Sub

    Private Sub txtDragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtDeSFolder.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub

End Class
