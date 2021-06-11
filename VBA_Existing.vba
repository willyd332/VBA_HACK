Option Base 1
Private Sub CommandButton_Click(Optional forceUpdate As Integer = 0)

           'Remove the filters
            ar = Array(Sheet3, Sheet4, Sheet5, Sheet8, Sheet9, Sheet10, Sheet11, Sheet12, Sheet14, Sheet19, Sheet21)

            For i = 1 To UBound(ar)
                If ar(i).FilterMode Then ar(i).ShowAllData
            Next i

   'Auto fill the formulas
    
     Dim isUpdate As Integer
     
     If forceUpdate = 1 Then
        isUpdate = vbYes
     Else
        isUpdate = MsgBox("Shall we update the projection?", vbYesNo, "Check before update")
     End If
     
     If isUpdate = vbYes Then
     
            Application.ScreenUpdating = False
            Application.Calculation = xlAutomatic
            
            Call Initial_var
            
            isize = Sheet3.Range("A11").End(xlDown).Row
            jsize = Sheet3.Range("K11").End(xlToRight).Column - 10
            
            Sheet8.Range("D4").Resize(isize - 11, 1).FormulaR1C1 = "=MAX(RC[-2], 0)"
            Sheet21.Range("D4").Resize(isize - 11, 1).FormulaR1C1 = "=MAX(RC[-2], 0)"
            
            Application.Calculation = xlManual
            



    
            Dim arr_activerate(), arr_failrate(), arr_bulletflag() As Variant
            Dim arr_pmt(), arr_actbal(), arr_totalsched() As Variant
            Dim arr_actbal_1(), arr_totalsched_1(), arr_regsched(), arr_bulletsched(), arr_pmt_ass(), arr_date() As Variant
        
            'pmt assumptions 1-first scheduled date 2- last sched month+1 3-Plan interest 4- payer group 5- inital adjustment 6-monbtly decay 9-Prepayment%
            arr_pmt_ass() = Sheet3.Range("B12").Resize(isize - 11, 9)
            arr_regsched() = Sheet11.Range("F4").Resize(isize - 11, jsize) 'regular scheduled
            arr_bulletsched() = Sheet12.Range("F4").Resize(isize - 11, jsize) 'bullet scheduled
            arr_date() = Sheet3.Range("K11").Resize(1, jsize) 'dates
        
    
            ReDim arr_activerate(1 To isize - 11, 1 To jsize)
            ReDim arr_failrate(1 To isize - 11, 1 To jsize)
            ReDim arr_bulletflag(1 To isize - 11, 1 To jsize)
            ReDim arr_servicefee(1 To isize - 11, 1 To jsize)
            ReDim arr_fixedcost(1 To isize - 11, 1 To jsize)
            
            bullet_rate = Sheet1.Cells(14, "B")
            servicefee = Sheet1.Cells(15, "B")
            vat_rate = Sheet1.Cells(16, "B")
         '------------------------------------------Active%------------arr_activerate ------------
            For i = 1 To isize - 11
                For j = 1 To jsize
                    If j = 1 Then
                        If arr_pmt_ass(i, 2) = "NA" Or arr_pmt_ass(i, 2) <= arr_date(1, j) Then
                            arr_activerate(i, j) = 0
                        ElseIf arr_bulletsched(i, j) > 0 Then
                            arr_activerate(i, j) = (1 - bullet_rate) * (1 - arr_pmt_ass(i, 6))
                        Else
                            arr_activerate(i, j) = arr_pmt_ass(i, 5) * (1 - arr_pmt_ass(i, 6))
                        End If
                    Else
                        If arr_pmt_ass(i, 2) <= arr_date(1, j) Then
                            arr_activerate(i, j) = 0
                        ElseIf arr_bulletsched(i, j) > 0 Then
                            arr_activerate(i, j) = arr_activerate(i, j - 1) * (1 - bullet_rate) * (1 - arr_pmt_ass(i, 6))
                        Else
                            arr_activerate(i, j) = arr_activerate(i, j - 1) * (1 - arr_pmt_ass(i, 6))
                        End If
                    End If
                Next j
            Next i
    
         '-----------------------------------------Failed%------------arr_activerate ------------
    
            For i = 1 To isize - 11
                For j = 1 To jsize
                    If j = 1 Then
                    
                        If arr_pmt_ass(i, 1) = "NA" Or arr_pmt_ass(i, 2) = "NA" Then
                            arr_failrate(i, j) = 1
                        ElseIf arr_pmt_ass(i, 1) > arr_date(1, j) Then
                            arr_failrate(i, j) = 0
                        Else
                            arr_failrate(i, j) = 1 - arr_activerate(i, j)
                        End If
                        
                    Else
                        If arr_pmt_ass(i, 1) = "NA" Or arr_pmt_ass(i, 2) = "NA" Or arr_pmt_ass(i, 2) <= arr_date(1, j) Then
                            arr_failrate(i, j) = 0
                        ElseIf arr_pmt_ass(i, 1) > arr_date(1, j) Then
                            arr_failrate(i, j) = 0
                        ElseIf arr_pmt_ass(i, 1) = arr_date(1, j) Then
                            arr_failrate(i, j) = 1 - arr_activerate(i, j)
                        Else
                            arr_failrate(i, j) = arr_activerate(i, j - 1) - arr_activerate(i, j)
                        End If
                        
                    End If
                Next j
            Next i
    
         '-----------------------------------------Bullet Flag------------arr_bulletflag ------------
            For i = 1 To isize - 11
                For j = 1 To jsize
                    If arr_bulletsched(i, j) > 0 Then
                        arr_bulletflag(i, j) = 1
                    Else
                        arr_bulletflag(i, j) = 0
                    End If
                    
                Next j
            Next i
            
         '-----------------------------------------Performing Loan Fee -------------------------------
            
       For i = 1 To isize - 11
                For j = 1 To jsize
                    If j = 1 Then
                       arr_fixedcost(i, j) = arr_fixedcost(i, j) + fc_payer * (1 + vat_rate) * arr_activebalanceMaxIsCool(i, j)
                    Else
                        arr_fixedcost(i, j) = arr_fixedcost(i, j) + fc_payer * (1 + vat_rate) * arr_activebalanceMaxIsCool(i, j)
                    End If
                Next j
            Next i
            
            
            
            
            
     
     If forceUpdate = 1 Then
        isUpdate = vbNo
     Else
        isUpdate = MsgBox("Shall we use autofill for the Excel formulas?", vbYesNo, "Check before update")
     End If
     
     If isUpdate = vbYes Then

            
                       'Clear the contents
            Sheet3.Range("K12").Resize(1000000, 16300).ClearContents
           Sheet8.Range("E4").Resize(1000000, 16300).ClearContents
            Sheet21.Range("E4").Resize(1000000, 16300).ClearContents
            Sheet5.Range("C4").Resize(1000000, 16300).ClearContents
            Sheet9.Range("C4").Resize(1000000, 16300).ClearContents
            Sheet10.Range("F4").Resize(1000000, 16300).ClearContents
                      
            ' Excel formula for the first line %active, %fail, bullet flag tabs
            Sheet10.Range("F4").Resize(isize - 11, jsize) = arr_activerate()
            Sheet10.Range("F4").FormulaR1C1 = "=IF(OR(PMT!R[8]C3=""NA"",PMT!R[8]C3<='Active %'!R3C),0,IF('Bullet Sched'!RC>0,(1-PriceAll!R14C2)*(1-RC4),RC3*(1-RC4)))"
            Sheet10.Range("G4").Resize(1, jsize - 1).FormulaR1C1 = "=IF(PMT!R[8]C3<='Active %'!R3C,0,IF('Bullet Sched'!RC>0,RC[-1]*(1-PriceAll!R14C2)*(1-RC4),RC[-1]*(1-RC4)))"
            
            Sheet9.Range("C4").Resize(isize - 11, jsize) = arr_failrate()
            Sheet9.Range("C4").FormulaR1C1 = "=IF(OR(PMT!R[8]C2=""NA"",PMT!R[8]C3=""NA""),100%,IF(PMT!R[8]C2>'Fail %'!R3C,0,1-'Active %'!RC[3]))"
            Sheet9.Range("D4").Resize(1, jsize - 1).FormulaR1C1 = "=IF(OR(PMT!R[8]C2=""NA"",PMT!R[8]C3=""NA"",PMT!R[8]C3<='Fail %'!R3C),0,IF(PMT!R[8]C2>'Fail %'!R3C,0,IF(PMT!R[8]C2='Fail %'!R3C,1-'Active %'!RC[3],'Active %'!RC[2]-'Active %'!RC[3])))"
            
            Sheet5.Range("C4").Resize(isize - 11, jsize) = arr_bulletflag()
            Sheet5.Range("C4").Resize(1, jsize).FormulaR1C1 = "=IF('Bullet Sched'!RC[3]>0,1,0)"
               
        
            arr_actbal_1() = Sheet8.Range("D4").Resize(isize - 11, 1) 'start actual balance
            arr_totalsched_1() = Sheet21.Range("D4").Resize(isize - 11, 1) 'total scheduled start
            
    
            ReDim arr_pmt(1 To isize - 11, 1 To jsize)
            ReDim arr_actbal(1 To isize - 11, 1 To jsize)
            ReDim arr_totalsched(1 To isize - 11, 1 To jsize)
            

            
            'On Error Resume Next
                
            For i = 1 To isize - 11
                For j = 1 To jsize
                    If j = 1 Then
                        arr_actbal(i, j) = arr_actbal_1(i, 1)
                        arr_totalsched(i, j) = arr_totalsched_1(i, 1)
                        arr_pmt(i, j) = WorksheetFunction.Min(arr_actbal(i, j), (arr_regsched(i, j) + arr_bulletsched(i, j) * bullet_rate + arr_actbal(i, j) * arr_pmt_ass(i, 9))) * arr_activerate(i, j)
                    Else
                        arr_actbal(i, j) = _
                            WorksheetFunction.Max(0, arr_actbal(i, j - 1) * (1 + arr_pmt_ass(i, 3) / 12) - WorksheetFunction.Min(arr_actbal(i, j - 1), (arr_actbal(i, j - 1) * arr_pmt_ass(i, 9) + arr_regsched(i, j - 1) + arr_bulletsched(i, j - 1) * bullet_rate)))
                        arr_totalsched(i, j) = WorksheetFunction.Max(0, arr_totalsched(i, j - 1) - arr_actbal(i, j - 1) * arr_pmt_ass(i, 9) - arr_regsched(i, j - 1) - arr_bulletsched(i, j - 1))
                        arr_pmt(i, j) = WorksheetFunction.Min(arr_actbal(i, j), (arr_regsched(i, j) + arr_bulletsched(i, j) * bullet_rate + arr_actbal(i, j) * arr_pmt_ass(i, 9))) * arr_activerate(i, j)
                    End If
                Next j
            Next i
            

            '-----------PMT-------------------
            Sheet3.Range("K12").Resize(isize - 11, jsize) = arr_pmt()
            Sheet3.Range("K12").Resize(1, jsize).FormulaR1C1 = _
                "=MIN('Active Bal'!R[-8]C[-7],('Reg Sched'!R[-8]C[-5]+'Active Bal'!R[-8]C[-7]*RC10+'Bullet Sched'!R[-8]C[-5]*(PriceAll!R14C2)))*'Active %'!R[-8]C[-5]"

            'Active Bal
            Sheet8.Range("D4").Resize(isize - 11, jsize) = arr_actbal()
            Sheet8.Range("D4").Resize(isize - 11, 1).FormulaR1C1 = "=MAX(RC[-2], 0)"
            Sheet8.Range("E4").Resize(1, jsize - 1).FormulaR1C1 = "=MAX(RC[-1]*(1+RC3/12)-MIN(RC[-1],(RC[-1]*PMT!R[8]C10+'Reg Sched'!RC[1]+'Bullet Sched'!RC[1]*PriceAll!R14C2)),0)"
            
            'Active Total Sched
            Sheet21.Range("D4").Resize(isize - 11, jsize) = arr_totalsched()
            Sheet21.Range("D4").Resize(isize - 11, 1).FormulaR1C1 = "=MAX(RC[-2], 0)"
            Sheet21.Range("E4").Resize(1, jsize - 1).FormulaR1C1 = "=MAX(RC[-1]-'Active Bal'!RC[-1]*PMT!R[8]C10-'Reg Sched'!RC[1]-'Bullet Sched'!RC[1], 0)"
            
'            Application.Calculation = xlAutomatic
'            Application.ScreenUpdating = True
            
    End If
    

'
'            Application.ScreenUpdating = False
'            Application.Calculation = xlManual
            
            Sheet4.Range("J12").Resize(1000000, 16300).ClearContents
            Sheet14.Range("J12").Resize(1000000, 16300).ClearContents
            Sheet19.Range("N12").Resize(1000000, 16300).ClearContents
            Sheet20.Range("N12").Resize(1000000, 16300).ClearContents
            Sheet26.Range("N12").Resize(1000000, 16300).ClearContents
                       

            Call VenteDPO_cal
            Call VenteForcee_cal
            Call newRCC_cal
   
            Application.Calculation = xlAutomatic
            Application.ScreenUpdating = True
            
            If forceUpdate <> 1 Then
                MsgBox "Done!", vbInformation
            End If
            
    End If
    
End Sub

Private Sub CommandButton1_Click()
    CommandButton_Click (0)
End Sub