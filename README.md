'# storemanager
'Store Management Application (Excel VBA)



Private Sub Adresaclient_Change()


If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 6) = Adresaclient.Text


End If



End Sub

Private Sub asteptarepiese_Click()



End Sub

Private Sub Buletinclient_Change()


If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 7) = Buletinclient.Text


End If





End Sub

Private Sub codpiesa1_Change()





End Sub





Private Sub cantitate1_Change()


            If cantitate1.Text = "" Then
            total1.Text = ""

End If


                If Not IsNumeric(cantitate1.Value) Then
                
                
                
                cantitate1.Text = ""
                
                End If

                
                If IsNumeric(pret1.Value) And IsNumeric(cantitate1.Value) Then
                
                total1.Text = pret1 * cantitate1
                
                
                End If
                
                
             Call detaliicomanda(2, 1)
             
             
             
             
            
             
         
             
              
              
             ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 23) = upgradenrprcerute
             
            
             
             
             
             




End Sub

Private Sub cantitate2_Change()


If cantitate2.Text = "" Then
total2.Text = ""

End If

If Not IsNumeric(cantitate2.Value) Then

'msgBox "Cantitatea trebuie sa fie numar!"

cantitate2.Text = ""

End If

If IsNumeric(pret2.Value) And IsNumeric(cantitate2.Value) Then

total2.Text = pret2 * cantitate2


End If



  Call detaliicomanda(2, 2)

Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 23) = upgradenrprcerute


End Sub

Private Sub cantitate3_Change()

        If cantitate3.Text = "" Then
        total3.Text = ""
        
        End If
        
        If Not IsNumeric(cantitate3.Value) Then
        
         'msgBox "Cantitatea trebuie sa fie numar!"
        
        cantitate3.Text = ""
        
        End If
        
        If IsNumeric(pret3.Value) And IsNumeric(cantitate3.Value) Then
        
        total3.Text = pret3 * cantitate3
        
        
        End If



  Call detaliicomanda(2, 3)

Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 23) = upgradenrprcerute

End Sub

Private Sub cantitate4_Change()


        If cantitate4.Text = "" Then
        total4.Text = ""

End If

If Not IsNumeric(cantitate4.Value) Then

'MsgBox "Cantitatea trebuie sa fie numar!"

cantitate4.Text = ""

End If

        If IsNumeric(pret4.Value) And IsNumeric(cantitate4.Value) Then
        
        total4.Text = pret4 * cantitate4


End If


                  Call detaliicomanda(2, 4)
        
        Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 23) = upgradenrprcerute
End Sub

Private Sub cantitate5_Change()


        If cantitate5.Text = "" Then
        total5.Text = ""

End If

If Not IsNumeric(cantitate5.Value) Then

    'MsgBox "Cantitatea trebuie sa fie numar!"
    
    cantitate5.Text = ""

End If


If IsNumeric(pret5.Value) And IsNumeric(cantitate5.Value) Then

total5.Text = pret5 * cantitate5


End If


   Call detaliicomanda(2, 5)
Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 23) = upgradenrprcerute


End Sub

            Private Sub cantitate6_Change()
            
            If cantitate6.Text = "" Then
            total6.Text = ""
            
            End If

            If Not IsNumeric(cantitate6.Value) Then
            
            'MsgBox "Cantitatea trebuie sa fie numar!"
            
            cantitate6.Text = ""
            
            End If

            If IsNumeric(pret6.Value) And IsNumeric(cantitate6.Value) Then
            
            total6.Text = pret6 * cantitate6
            
            
            End If


     Call detaliicomanda(2, 6)
Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 23) = upgradenrprcerute

End Sub



Private Sub cantitatemo2_Change()




End Sub

Private Sub cantitate7_Change()


            If cantitate7.Text = "" Then
    
            total7.Text = ""
            
            End If

            If Not IsNumeric(cantitate7.Value) Then
            'MsgBox "Cantitatea trebuie sa fie numar!"
            
            cantitate7.Text = ""
            
            End If

            If IsNumeric(pret7.Value) And IsNumeric(cantitate7.Value) Then
            
            total7.Text = pret7 * cantitate7
            
            
            End If








 Call detaliicomanda(2, 7)



End Sub

Private Sub cantitate8_Change()


             If cantitate8.Text = "" Then
    
            total8.Text = ""
            
            End If

            If Not IsNumeric(cantitate8.Value) Then
            
             'msgBox "Cantitatea trebuie sa fie numar!"
            
            cantitate8.Text = ""
            
            End If

            If IsNumeric(pret8.Value) And IsNumeric(cantitate8.Value) Then
            
            total8.Text = pret8 * cantitate8
            
            
            End If





 Call detaliicomanda(2, 8)



End Sub


Private Sub cantitate9_Change()

    If cantitate9.Text = "" Then
    
            total9.Text = ""
            
            End If

            If Not IsNumeric(cantitate9.Value) Then
            
            'msgBox "Cantitatea trebuie sa fie numar!"
            
            cantitate9.Text = ""
            
            End If

            If IsNumeric(pret9.Value) And IsNumeric(cantitate9.Value) Then
            
            total9.Text = pret9 * cantitate9
            
            
            End If




 Call detaliicomanda(2, 9)



End Sub

Private Sub cod1_Change()





            denumire1.Locked = False
            denumire1.Text = ""

            
            If Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod1.Text) > 0 Then
            
            randuldenumire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod1.Text) + 1
            
            
            
            denumire1.Text = ThisWorkbook.Sheets("intrare piese").Cells(randuldenumire, 5)
            
            denumire1.Locked = True






                End If
                
             Call detaliicomanda(1, 1)

            




End Sub

Private Sub cod2_Change()




denumire2.Locked = False
denumire2.Text = ""


If Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod2.Text) > 0 Then

randuldenumire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod2.Text) + 1



denumire2.Text = ThisWorkbook.Sheets("intrare piese").Cells(randuldenumire, 5)


denumire2.Locked = True


End If

         Call detaliicomanda(1, 2)








End Sub

Private Sub cod3_Change()

'MsgBox "cod3"

denumire3.Locked = False
denumire3.Text = ""


If Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod3.Text) > 0 Then

randuldenumire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod3.Text) + 1



denumire3.Text = ThisWorkbook.Sheets("intrare piese").Cells(randuldenumire, 5)

denumire3.Locked = True

End If

Call detaliicomanda(1, 3)




End Sub

Private Sub cod4_Change()

denumire4.Locked = False
denumire4.Text = ""


If Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod4.Text) > 0 Then

randuldenumire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod4.Text) + 1



denumire4.Text = ThisWorkbook.Sheets("intrare piese").Cells(randuldenumire, 5)

denumire4.Locked = True

End If


Call detaliicomanda(1, 4)








End Sub

Private Sub cod5_Change()



denumire5.Locked = False
denumire5.Text = ""


If Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod5.Text) > 0 Then

randuldenumire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod5.Text) + 1



denumire5.Text = ThisWorkbook.Sheets("intrare piese").Cells(randuldenumire, 5)

denumire5.Locked = True

End If

Call detaliicomanda(1, 5)




End Sub

Private Sub cod6_Change()

denumire6.Locked = False
denumire6.Text = ""


If Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod6.Text) > 0 Then

randuldenumire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("intrare piese").Range("b:b"), ThisWorkbook.Sheets("intrare piese").Range("c:c"), cod6.Text) + 1



denumire6.Text = ThisWorkbook.Sheets("intrare piese").Cells(randuldenumire, 5)

denumire6.Locked = True

End If

Call detaliicomanda(1, 6)




End Sub

Private Sub cod7_Change()


Call detaliicomanda(1, 7)



End Sub
Private Sub cod8_Change()


Call detaliicomanda(1, 8)



End Sub

Private Sub cod9_Change()


Call detaliicomanda(1, 9)



End Sub

Private Sub codpiesa_Click()

End Sub

Private Sub Comandanoua_Click()

               
                
              
            '  ThisWorkbook.Sheets("selectie numeclienti").Cells(1, 12) = ""
               


                   
                            

                            User.Caption = ThisWorkbook.Sheets("Start aplicatie").Cells(4, 2)
    

                            User.Visible = True
                        
                            nouacomanda.Visible = True
                            
                            Validarecomanda.Visible = True
                            stergerecomanda.Visible = True
                            Duplicarecomanda.Visible = True
                             facturarecomanda.Visible = True
                
                            statutcomanda.Visible = True
                            
                            
                            marciauto.Visible = True
                            Modeleauto.Visible = True
                    
                
                            Listapiesa.Visible = True
                            
                            nrpiesa.Visible = True
                            codpiesa.Visible = True
                            denumirepiesa.Visible = True
                            cantitatepiesa.Visible = True
                            totallinie.Visible = True
                            pretvanzarepiesa.Visible = True
                            
                            piesesosite.Visible = True
                            piesesosite1.Visible = True
                            piesesosite2.Visible = True
                            piesesosite3.Visible = True
                            piesesosite4.Visible = True
                            piesesosite5.Visible = True
                            piesesosite6.Visible = True
            
            
            
            
            
            
            
            nr1.Visible = True
            cod1.Visible = True
            cantitate1.Visible = True
            pret1.Visible = True
            denumire1.Visible = True
            total1.Visible = True
            pret1.Visible = True
            
            nr2.Visible = True
            cod2.Visible = True
            cantitate2.Visible = True
            pret2.Visible = True
            denumire2.Visible = True
            total2.Visible = True
            pret2.Visible = True
            
            nr3.Visible = True
            cod3.Visible = True
            cantitate3.Visible = True
            pret3.Visible = True
            denumire3.Visible = True
            total3.Visible = True
            pret3.Visible = True
            
            nr4.Visible = True
            cod4.Visible = True
            cantitate4.Visible = True
            pret4.Visible = True
            denumire4.Visible = True
            total4.Visible = True
            pret4.Visible = True
            
            nr5.Visible = True
            cod5.Visible = True
            cantitate5.Visible = True
            pret5.Visible = True
            denumire5.Visible = True
            total5.Visible = True
            pret5.Visible = True
            
            
            nr6.Visible = True
            cod6.Visible = True
            cantitate6.Visible = True
            pret6.Visible = True
            denumire6.Visible = True
            total6.Visible = True
            pret6.Visible = True
            
             nr9.Visible = True
            cod9.Visible = True
            cantitate9.Visible = True
            pret9.Visible = True
            denumire9.Visible = True
            total9.Visible = True
            pret9.Visible = True
            
            
             nr7.Visible = True
            cod7.Visible = True
            cantitate7.Visible = True
            pret7.Visible = True
            denumire7.Visible = True
            total7.Visible = True
            pret7.Visible = True
            
            
            
             nr8.Visible = True
            cod8.Visible = True
            cantitate8.Visible = True
            pret8.Visible = True
            denumire8.Visible = True
            total8.Visible = True
            pret8.Visible = True
            
            
            totalpr.Visible = True
            totalmo.Visible = True
            totaltvapr.Visible = True
            totaltvamo.Visible = True
            totalcda.Visible = True
            totalcdatva.Visible = True
            totalplata.Visible = True
            


            


        nr1.Text = 1
        nr2.Text = 2
        nr3.Text = 3
        nr4.Text = 4
        nr5.Text = 5
        nr6.Text = 6
        nr7.Text = 7
        nr8.Text = 8
        nr9.Text = 9



                Numeclient.Visible = True
                Prenumeclient.Visible = True
                Telefonclient.Visible = True
                Buletinclient.Visible = True
                Adresaclient.Visible = True
                Numarcomanda.Visible = True
                Numar.Visible = True
                Marcamasina.Visible = True
                Modelmasina.Visible = True
                Kilometraj.Visible = True
                Marca.Visible = True
                Model.Visible = True
                Rulaj.Visible = True
                Masinainmatriculare.Visible = True
                Inmatriculare.Visible = True
                
                
                
                
                Nume.Visible = True
                Prenume.Visible = True
                Telefon.Visible = True
                Buletin.Visible = True
                Adresa.Visible = True
                DDR.Visible = True
                Datadeschidere.Visible = True
                datafacturare.Visible = True
                datafacturarii.Visible = True
                Datalivrare.Visible = True
                datalivrarii.Visible = True
                
                
                
                
                
                
                    
        Numarcomanda.Text = 1
                    
         '   MsgBox Numarcomanda.Text
                    
                    
                    
                  



    
               
                    
                    
                    




End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub comandapartiala_Click()




End Sub

Private Sub DDR_Change()





 
        
        
       
        
        

End Sub

Private Sub Delivrat_Click()

End Sub

Private Sub denumire1_Change()


 Call detaliicomanda(3, 1)








End Sub
Private Sub denumire2_Change()


Call detaliicomanda(3, 2)








End Sub

Private Sub denumire3_Change()



 Call detaliicomanda(3, 3)
                
                

End Sub

Private Sub denumire4_Change()


 Call detaliicomanda(3, 4)






End Sub

Private Sub denumire5_Change()

 Call detaliicomanda(3, 5)








End Sub

Private Sub denumire6_Change()




Call detaliicomanda(3, 6)












End Sub

Private Sub denumire7_Change()


Call detaliicomanda(3, 7)



End Sub

Private Sub denumire8_Change()


Call detaliicomanda(3, 8)



End Sub


Private Sub denumire9_Change()


Call detaliicomanda(3, 9)



End Sub

Private Sub Duplicarecomanda_Click()


For h = 1 To 9

resetarepiese (h)

Next h







'MsgBox cdavalidata

                        If cdavalidata = 1 Then



                            

                        x = Numarcomanda.Value + 1
                        
                      '  MsgBox "x=" & x
                      
                        ver = ThisWorkbook.Sheets("Comenzisimplu").Cells(1, 1)
                        
                       ' MsgBox "ver=" & ThisWorkbook.Sheets("Comenzisimplu").Cells(ver + 1, 14)
                        
                        If ThisWorkbook.Sheets("Comenzisimplu").Cells(ver + 1, 14) <> 1 Then
                        
                                    Numarcomanda.Value = ThisWorkbook.Sheets("Comenzisimplu").Cells(1, 1)
                                    
                                    Else: Numarcomanda.Value = ThisWorkbook.Sheets("Comenzisimplu").Cells(1, 1) + 1
                        
                        End If
                        
                        
                        
                        
                        
                        
                                        Numeclient.Text = ""
                                     Prenumeclient.Text = ""
                                     Telefonclient.Text = ""
                                     
                                     Adresaclient.Text = ""
                                     Buletinclient.Text = ""
                                    
                                     
                                     Marcamasina.Text = ""
                                     Modelmasina.Text = ""
                                     Masinainmatriculare.Text = ""
                                     Kilometraj.Text = ""
                                   
                                     
                                    
                                    
                                     
                                     
                                     
                                  '   nr1.Text = ""
                                     cod1.Text = ""
                                     denumire1.Text = ""
                                     cantitate1.Text = ""
                                     pret1.Text = ""
                                     total1.Text = ""
                                     piesesosite1.Text = ""
                                     
                                    '  nr2.Text = ""
                                     cod2.Text = ""
                                     denumire2.Text = ""
                                     cantitate2.Text = ""
                                     pret2.Text = ""
                                     total2.Text = ""
                                     piesesosite2.Text = ""
                                     
                                    '  nr3.Text = ""
                                     cod3.Text = ""
                                     denumire3.Text = ""
                                     cantitate3.Text = ""
                                     pret3.Text = ""
                                     total3.Text = ""
                                     piesesosite3.Text = ""
                                     
                                     
                                  '    nr4.Text = ""
                                     cod4.Text = ""
                                     denumire4.Text = ""
                                     cantitate4.Text = ""
                                     pret4.Text = ""
                                     total4.Text = ""
                                     piesesosite4.Text = ""
                                     
                                     
                                    '  nr5.Text = ""
                                     cod5.Text = ""
                                     denumire5.Text = ""
                                     cantitate5.Text = ""
                                     pret5.Text = ""
                                     total5.Text = ""
                                     piesesosite5.Text = ""
                                     
                                     
                                  '    nr6.Text = ""
                                     cod6.Text = ""
                                     denumire6.Text = ""
                                     cantitate6.Text = ""
                                     pret6.Text = ""
                                     total6.Text = ""
                                     piesesosite6.Text = ""
                                     
                                     
                                 '    nr7.Text = ""
                                     cod7.Text = ""
                                     denumire7.Text = ""
                                     cantitate7.Text = ""
                                     pret7.Text = ""
                                     total7.Text = ""
                                     piesesosite7.Text = ""
                                     
                                     
                                  '   nr8.Text = ""
                                     cod8.Text = ""
                                     denumire8.Text = ""
                                     cantitate8.Text = ""
                                     pret8.Text = ""
                                     total8.Text = ""
                                     piesesosite8.Text = ""
                                     
                                   '  nr9.Text = ""
                                     cod9.Text = ""
                                     denumire9.Text = ""
                                     cantitate9.Text = ""
                                     pret9.Text = ""
                                     total9.Text = ""
                                     piesesosite9.Text = ""
                        
                        






                        'actualizare data deschidere in sheetul "comenzi simplu"
                         'actualizare data deschidere in sheet si in campul DDR
                                            
                                            Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 19) = CDate(Now())
                                            
                                            
                                            DDR.Value = Now()
                                            
                                            an = Year(DDR.Value)
                                            luna = Month(DDR.Value)
                                            zi = Day(DDR.Value)
                                            
                                            ddrafisat = zi & "/" & luna & "/" & an
                                            
                                            DDR.Text = ddrafisat
                    






                                     Numeclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 3)
                                     Prenumeclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 4)
                                     Telefonclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 5)
                                     
                                     Adresaclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 6)
                                     Buletinclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 7)
                                    
                                     
                                     Marcamasina.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 8)
                                     Modelmasina.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 9)
                                     Masinainmatriculare.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 11)
                                     Kilometraj.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 10)
                                     marciauto.Text = Marcamasina.Text
                                     Modeleauto.Text = Modelmasina.Text
                                     
                                     'completarea liniilor de piese
                                     
                                     For lin = 1 To 9
                                     
                                     
                                     
                                     z = 1 + Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("Comenzi detaliate").Range("b:b"), ThisWorkbook.Sheets("Comenzi detaliate").Range("c:c"), x - 1, ThisWorkbook.Sheets("Comenzi detaliate").Range("d:d"), lin)
                                     
                                     
                                    ' MsgBox "z=" & z
                                    
                                    
                                        If z > 1 Then
                                        
                                      inutil = preluarepiese(z, lin)
                                      
                                      
                                    
                                      
                                 
                                     End If
                                     
                                     
                                     
                                     
                                     If Me.Controls("piesesosite" & lin).Text <> "" Then
                                     
                                     Me.Controls("piesesosite" & lin).Text = ""
                                     
                                     
                                     End If
                                     
                                     
                                     
                                     
                                     Next lin
                                     
                                     'sfr  completarea liniilor de piese
                                     
                                     
                                     
                                     
                                     
                                     


Else: MsgBox "Nu se poate duplica o comanda partiala"


End If







End Sub

Function cdavalidata()

If ThisWorkbook.Sheets("Comenzisimplu").Cells(Numarcomanda.Value + 1, 14) = 1 Then
cdavalidata = 1

Else: cdavalidata = 0


End If


End Function

Function cdadelivrat()

If ThisWorkbook.Sheets("Comenzisimplu").Cells(Numarcomanda.Value + 1, 15) = 1 Then
cdadelivrat = 1

Else: cdadelivrat = 0


End If


End Function

Function cdalivrata()

If ThisWorkbook.Sheets("Comenzisimplu").Cells(Numarcomanda.Value + 1, 16) = 1 Then
cdalivrata = 1

Else: cdalivrata = 0


End If


End Function

Function cdafacturata()

If ThisWorkbook.Sheets("Comenzisimplu").Cells(Numarcomanda.Value + 1, 17) = 1 Then
cdafacturata = 1

Else: cdafacturata = 0


End If


End Function

Function cdastearsa()

If ThisWorkbook.Sheets("Comenzisimplu").Cells(Numarcomanda.Value + 1, 25) = 1 Then
cdastearsa = 1

Else: cdastearsa = 0


End If


End Function

Private Sub facturarecomanda_Click()


 If cdastearsa = 0 And cdadelivrat = 1 And cdafacturata = 0 Then
 
 
            Facturata.Value = True
            
            
            mod_statut = ThisWorkbook.Sheets("modificare statut comenzi").Cells(1, 1) + 2
            ThisWorkbook.Sheets("modificare statut comenzi").Cells(mod_statut, 2) = mod_statut - 1
            ThisWorkbook.Sheets("modificare statut comenzi").Cells(mod_statut, 3) = Numarcomanda.Value
            ThisWorkbook.Sheets("modificare statut comenzi").Cells(mod_statut, 4) = "Comanda facturata"
            ThisWorkbook.Sheets("modificare statut comenzi").Cells(mod_statut, 5) = CDate(Now())
            ThisWorkbook.Sheets("modificare statut comenzi").Cells(mod_statut, 6) = User.Caption
            
            
            ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 17) = 1
            ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 22) = CDate(Now())
            ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 26) = User.Caption
            
            
            
            
            nr_linii_factura = Application.WorksheetFunction.CountIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), Numarcomanda.Value)
            
            
            
 
     
            
            
           
            
            
            contor_linii_facturi = ThisWorkbook.Sheets("lista facturi").Cells(1, 1) + 2
            
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 2) = contor_linii_facturi - 1
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 3) = contor_linii_facturi - 1
            
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 4) = Numeclient.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 5) = Prenumeclient.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 6) = Adresaclient.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 7) = Buletinclient.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 8) = Marcamasina.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 9) = Modelmasina.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 10) = Masinainmatriculare.Text
            
            
            For j = 1 To 9
            
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 10 + j) = Me.Controls("cod" & j).Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 19 + j) = Me.Controls("cantitate" & j).Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 28 + j) = Me.Controls("denumire" & j).Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 37 + j) = Me.Controls("pret" & j).Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 46 + j) = Me.Controls("total" & j).Text
            
            
            
            
            If IsNumeric(Me.Controls("pret" & j).Value) Then
            
                        Dim pretcutva As Double
                        
                        pretcutva = Me.Controls("total" & j).Value
                        
                        Dim valtva As Double
                        valtva = (ThisWorkbook.Sheets("intrare piese").Cells(1, 15)) / 100
                        
                        MsgBox " pret cu tva=" & pretcutva
                        MsgBox "  tva=" & valtva
                        
                        
                        ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 55 + j) = pretcutva * valtva
            
            End If
            
            
            Next j
            
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 65) = totalpr.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 66) = totalmo.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 67) = totaltvapr.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 68) = totaltvamo.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 69) = totalcda.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 70) = totalcdatva.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 71) = totalplata.Text
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 72) = CDate(Now())
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 73) = User.Caption
            ThisWorkbook.Sheets("lista facturi").Cells(contor_linii_facturi, 74) = Numarcomanda.Value


End If



End Sub

Private Sub Facturata_Click()

End Sub

Private Sub Inmatriculare_Click()

End Sub



Private Sub intrarepiese_Click()


Aplicatie.Hide

intraridepiese.Show vbModeless





End Sub

Private Sub Kilometraj_Change()


If Not IsNumeric(Kilometraj.Value) Then



Kilometraj.Text = ""


End If



If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 10) = Kilometraj.Text


End If



End Sub

Private Sub Label1_Click()

End Sub

Private Sub Listapiese_Click()

End Sub

Private Sub Livrata_Click()

End Sub

Private Sub Marcamasina_Change()


        ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 8) = Marcamasina.Text

        criteriulauto = Marcamasina.Text


    ThisWorkbook.Sheets("lista auto").UsedRange.AutoFilter _
    field:=1, _
    Criteria1:=criteriulauto, _
    VisibleDropDown:=False
 
      ThisWorkbook.Sheets("lista auto filtre").UsedRange.Clear
  
      ThisWorkbook.Sheets("lista auto").UsedRange.Copy Destination:=ThisWorkbook.Sheets("lista auto filtre").Range("a1")
      
      


            If ThisWorkbook.Sheets("lista auto").AutoFilterMode Then
             
                ThisWorkbook.Sheets("lista auto").AutoFilterMode = False
             
            End If
            
            Modelmasina.Text = ThisWorkbook.Sheets("lista auto filtre").Cells(2, 2)
            




End Sub

Private Sub marciauto_Change()


Marcamasina.Text = marciauto.Text





End Sub

Private Sub Masinainmatriculare_Change()


If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 11) = Masinainmatriculare.Text


End If



End Sub

Private Sub Model_Click()

End Sub

Private Sub Modeleauto_Change()


Modelmasina.Text = Modeleauto.Text



End Sub

Private Sub Modelmasina_Change()

If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 9) = Modelmasina.Text


End If




End Sub

Private Sub nouacomanda_Click()



Numarcomanda.Value = ThisWorkbook.Sheets("comenzisimplu").Cells(1, 1) + 1







End Sub

Private Sub nr7_Change()




End Sub

Private Sub Numar_Click()

End Sub

Private Sub pret7_Change()



                    If pret7.Text = "" Then
                    total7.Text = ""
                    
                    
                    
                    End If
                    


If Not IsNumeric(pret7.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret7.Text = ""

End If

If IsNumeric(pret7.Value) And IsNumeric(cantitate7.Value) Then

total7.Text = pret7 * cantitate7


End If






Call detaliicomanda(4, 7)



End Sub

Private Sub pret8_Change()



                    If pret8.Text = "" Then
                    total8.Text = ""
                    
                    
                    
                    End If
                    


If Not IsNumeric(pret8.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret8.Text = ""

End If

If IsNumeric(pret8.Value) And IsNumeric(cantitate8.Value) Then

total8.Text = pret8 * cantitate8


End If


Call detaliicomanda(4, 8)



End Sub

Private Sub pret9_Change()


                    If pret9.Text = "" Then
                    total9.Text = ""
                    
                    
                    
                    End If
                    

                    
                    If Not IsNumeric(pret9.Value) Then
                    
                    'MsgBox "Pretul trebuie sa fie numeric!"
                    
                    pret9.Text = ""
                    
                    End If
                    
                    If IsNumeric(pret9.Value) And IsNumeric(cantitate9.Value) Then
                    
                    total9.Text = pret9 * cantitate9
                    
                    
                    End If





Call detaliicomanda(4, 9)



End Sub

Private Sub stergerecomanda_Click()



    
    
    'doar daca dt e sters
    
    If cdastearsa = 0 Then
    
    
            
                
                'daca devizul este partial se elimina toate campurile completate fara a fi si stearsa
                
                
                If cdavalidata = 0 Then
    
    


                                     Numeclient.Text = ""
                                     Prenumeclient.Text = ""
                                     Telefonclient.Text = ""
                                     
                                     Adresaclient.Text = ""
                                     Buletinclient.Text = ""
                                    
                                     
                                     Marcamasina.Text = ""
                                     Modelmasina.Text = ""
                                     Masinainmatriculare.Text = ""
                                     Kilometraj.Text = ""
                                   
                                     
                                    
                                    
                                     
                                     
                                     
                                   ''  nr1.Text = ""
                                     cod1.Text = ""
                                     denumire1.Text = ""
                                     cantitate1.Text = ""
                                     pret1.Text = ""
                                     total1.Text = ""
                                     piesesosite1.Text = ""
                                     
                                   ''   nr2.Text = ""
                                     cod2.Text = ""
                                     denumire2.Text = ""
                                     cantitate2.Text = ""
                                     pret2.Text = ""
                                     total2.Text = ""
                                     piesesosite2.Text = ""
                                     
                                    ''  nr3.Text = ""
                                     cod3.Text = ""
                                     denumire3.Text = ""
                                     cantitate3.Text = ""
                                     pret3.Text = ""
                                     total3.Text = ""
                                     piesesosite3.Text = ""
                                     
                                     
                                 ''     nr4.Text = ""
                                     cod4.Text = ""
                                     denumire4.Text = ""
                                     cantitate4.Text = ""
                                     pret4.Text = ""
                                     total4.Text = ""
                                     piesesosite4.Text = ""
                                     
                                     
                                   ''   nr5.Text = ""
                                     cod5.Text = ""
                                     denumire5.Text = ""
                                     cantitate5.Text = ""
                                     pret5.Text = ""
                                     total5.Text = ""
                                     piesesosite5.Text = ""
                                     
                                     
                                  ''    nr6.Text = ""
                                     cod6.Text = ""
                                     denumire6.Text = ""
                                     cantitate6.Text = ""
                                     pret6.Text = ""
                                     total6.Text = ""
                                     piesesosite6.Text = ""
                                     
                                     
                                   ''  nr7.Text = ""
                                     cod7.Text = ""
                                     denumire7.Text = ""
                                     cantitate7.Text = ""
                                     pret7.Text = ""
                                     total7.Text = ""
                                     piesesosite7.Text = ""
                                     
                                     
                                   ''  nr8.Text = ""
                                     cod8.Text = ""
                                     denumire8.Text = ""
                                     cantitate8.Text = ""
                                     pret8.Text = ""
                                     total8.Text = ""
                                     piesesosite8.Text = ""
                                     
                                  ''   nr9.Text = ""
                                     cod9.Text = ""
                                     denumire9.Text = ""
                                     cantitate9.Text = ""
                                     pret9.Text = ""
                                     total9.Text = ""
                                     piesesosite9.Text = ""



              End If
    
    
        If cdavalidata = 1 Then
        
        MsgBox "Comanda a fost stearsa!"
        stearsa.Value = True
        
        
        'stergere doar dc nu e partiala
        
        
        
        ThisWorkbook.Sheets("Comenzisimplu").Cells(Numarcomanda.Value + 1, 25) = 1
        
        
        For x = 1 To 6
        
          If Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("stoc piese").Range("C:c"), Me.Controls("cod" & x).Text, ThisWorkbook.Sheets("stoc piese").Range("b:b")) > 0 Then
        
                MsgBox "pr sosite=" & Me.Controls("piesesosite" & x).Value
        
                If Me.Controls("cantitate" & x).Text <> "" And Me.Controls("cantitate" & x).Value <> 0 Then
                
                
                                    pr_intrare = Me.Controls("cod" & x).Text
                                    qnt_intrare = Me.Controls("cantitate" & x).Value
                                    
                                    MsgBox "piese intrare=" & pr_intrare
                                       'introducere detalii in intrare piese
                                     
                                     linie_adaugare = Sheets("intrare piese").Cells(1, 1) + 2
                                     
                                     
                                     
                                     'Dim pret_vi As Double
                                     
                                    ' pret_vi = Application.WorksheetFunction.Round(Me.Controls("pret" & x).Value, 2)
                                     
                                     
                                     
                                     MsgBox "linie=" & linie_adaugare
                                     
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 2) = ThisWorkbook.Sheets("intrare piese").Cells(1, 1) + 1
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 3) = pr_intrare
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 4) = Me.Controls("denumire" & x).Text
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 5) = qnt_intrare
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 10) = Numarcomanda.Value
                                    ' thisworkbook.sheets("intrare piese").Cells(linie_adaugare, 7) = pretachizitieintrare.Value
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 8) = Me.Controls("pret" & x).Text
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 11) = CDate(Now())
                                     ThisWorkbook.Sheets("intrare piese").Cells(linie_adaugare, 17) = User.Caption
                                     Me.Controls("piesesosite" & x) = ""
                                     
                                     
                                     
                                     ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 24) = 0
                                     
                                      'sfr introducere detalii in intrare piese
                                    
                                    
                                    'actualizare sc
                                    
                                       
                                
                         
                                
                        'daca exista reperul se updateaza stocul
                        If Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("stoc piese").Range("C:c"), pr_intrare, ThisWorkbook.Sheets("stoc piese").Range("b:b")) > 0 Then
                
                              
                                     
                                     
                                        Set cautare_pr_stoc = ThisWorkbook.Sheets("stoc piese").Range("c1")
                                        Set cautare_pr_stoc = ThisWorkbook.Sheets("stoc piese").Columns(3).Find(What:=pr_intrare, After:=cautare_pr_stoc, _
                                        LookIn:=xlValues, LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                         MatchCase:=True)
                                    
                                        
                                     
                                        
                                        
                                    rand_pr_stoc = cautare_pr_stoc.Row
                                        
                                        
                                    ThisWorkbook.Sheets("stoc piese").Cells(rand_pr_stoc, 4) = ThisWorkbook.Sheets("stoc piese").Cells(rand_pr_stoc, 4) + qnt_intrare
                                                      
                                    
                              End If
                                    
                                    'sfr actualizare stc
                                    
                                    
                                    
                                    
                                    
                                    'actualizare c_det
                                    
        
                                
                                            Dim iLoop2 As Integer
        

                                            Dim cautare_dt_det As Range
                                            Set cautare_dt_det = ThisWorkbook.Sheets("comenzi detaliate").Range("c1")
                                           
                
                
                                         iLoop2 = WorksheetFunction.CountIf(ThisWorkbook.Sheets("comenzi detaliate").Columns(3), Numarcomanda.Value)
                                         
                                         MsgBox "iloop2=" & iLoop2
                                         
                                         For j = 1 To iLoop2
                                    
                                       
                                        Set cautare_dt_det = ThisWorkbook.Sheets("comenzi detaliate").Columns(3).Find(What:=Numarcomanda.Value, After:=cautare_dt_det, _
                                        LookIn:=xlValues, LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                         MatchCase:=True)
                                    
                                        
                                        
                                        
                                        
                                        rand_dt_det = cautare_dt_det.Row
                                        
                                        MsgBox "rand cda det=" & rand_dt_det
                                        
                                        If rand_dt_det = x + 1 Then
                                        
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 10) = 0
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 18) = ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 6)
                                        
                                        
                                        End If
                                        
                                        
                                    
                                  
                                          Next j
                                    
                                    
                                 
                                   'sfr actualizare c_det
                                   
                                   
                    'cautare comenzi validate care solicita pr intrata




        MsgBox "aici"
        
        
        Dim nrcdagasita As Integer
        
        valstoc = ThisWorkbook.Sheets("stoc piese").Cells(rand_pr_stoc, 4)
        
        MsgBox "valstoc=" & valstoc
        
        referintapr = pr_intrare
        
        MsgBox "referinta pr=" & referintapr
        
                Dim iLoop3 As Integer
        

                Dim rNa3 As Range
                
               
                
                
                
                iLoop3 = WorksheetFunction.CountIf(ThisWorkbook.Sheets("comenzi detaliate").Columns(5), referintapr)
                
                Set rNa3 = ThisWorkbook.Sheets("comenzi detaliate").Range("e1")
                
                MsgBox "iloop=" & iLoop3
                
                
                
                
                 For i = 1 To iLoop3
                 
                 MsgBox "referinta pr=" & referintapr
                
                    Set rNa3 = ThisWorkbook.Sheets("comenzi detaliate").Columns(5).Find(What:=referintapr, After:=rNa3, _
                    LookIn:=xlValues, LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                     MatchCase:=True)
                
                  
                      
                      
                      
                      randulprgasita = rNa3.Row
                      
                     
                      
                      
                      
                      
                      
                      
                      nrcdaprgasita = ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 3)
                       MsgBox " nr cda  gasita=" & nrcdaprgasita
                      
                     
                      
                      
                      statutcdaprgasita = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), nrcdaprgasita, ThisWorkbook.Sheets("comenzisimplu").Range("n:n"))
                      statutstearsa = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), nrcdaprgasita, ThisWorkbook.Sheets("comenzisimplu").Range("y:y"))
                      
                     ' MsgBox nrcdaprgasita & " " & statutcdaprgasita
                      
                      'actualizare stoc, rest de sosit, iesiri daca dt este validat
                      
                      
                      If statutcdaprgasita = 1 And statutstearsa <> 1 Then
                      
                      
                      ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 18) = ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 6) - ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 10)
                      
                      
                      restsosire = ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 18)
                      
                      MsgBox "restsosire=" & restsosire
                      
                      maxintrare = Application.WorksheetFunction.Min(restsosire, valstoc)
                      
                      
                      
                                         MsgBox "nr cda1=" & nrcdaprgasita
                      
                                        If restsosire > 0 Then
                                        
                                        
                                        MsgBox "res >0 nr cda1=" & nrcdaprgasita
                                        
                                        restsosire = restsosire - maxintrare
                                        
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 18) = restsosire
                                        
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 10) = maxintrare + ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 10)
                                        
                                        valstoc = valstoc - maxintrare
                                        
                                        
                                        liniestoc = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("stoc piese").Range("c:c"), referintapr, ThisWorkbook.Sheets("stoc piese").Range("b:b")) + 1
                                        
                                        ThisWorkbook.Sheets("stoc piese").Cells(liniestoc, 4) = valstoc
                                        
                                        
                                         'actualizare iesire piese
                                         
                                         linieiesirepiese = ThisWorkbook.Sheets("iesire piese").Cells(1, 1) + 2
                                        
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 2) = ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese - 1, 2) + 1
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 3) = referintapr
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 4) = maxintrare
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 5) = Me.Controls("denumire" & x).Text
                                         
                                         'nume client
                                         
                                         poznumeclient = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), nrcdaprgasita, ThisWorkbook.Sheets("comenzisimplu").Range("b:b"))
                                         
                                         
                                         
                                         MsgBox "poz nume client=" & poznumeclient
                                         MsgBox "nr comanda=" & nrcdaprgasita
                                         
                                         
                                           Dim rangeclient As Range
                                         
                                           Set rangeclient = ThisWorkbook.Sheets("comenzisimplu").Range("B1")
                                         
                                            Set rangeclient = Columns(2).Find(What:=nrcdaprgasita, After:=rangeclient, _
                                            LookIn:=xlValues, LookAt:=xlWhole, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                            MatchCase:=True)
                                            
                                            
                                         
                                         
                                         MsgBox "range client row=" & rangeclient.Row
                                         
                                         Dim pretvanzintrare1 As Double
                                         pretvanzintrare1 = Me.Controls("pret" & x).Value
                                         
                                         
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 6) = ThisWorkbook.Sheets("comenzisimplu").Cells(rangeclient.Row, 3)
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 7) = ThisWorkbook.Sheets("comenzisimplu").Cells(rangeclient.Row, 4)
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 8) = pretvanzintrare1
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 10) = nrcdaprgasita
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 11) = CDate(Now())
                                            
                                            
                                            
                                        
                                        End If
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      End If
                      
      


             'actualizare stoc, rest de sosit, iesiri daca dt este validat


        If valstoc = 0 Then
        i = iLoop3 + 1
        End If
        
        
        
        'actualizare statut comanda de livrat
        
        nr_pr_cerute_cda = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), nrcdaprgasita, ThisWorkbook.Sheets("comenzi detaliate").Range("f:f"))
        nr_pr_sosite_cda = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), nrcdaprgasita, ThisWorkbook.Sheets("comenzi detaliate").Range("j:j"))
                
                
                
                If nr_pr_cerute_cda = nr_pr_sosite_cda And ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 15) <> 1 And ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 25) <> 1 Then
                
                MsgBox "comanda " & nrcdaprgasita & " este de livrat!"
                
                
                ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 15) = 1
                
                
                
                
                End If
                
                
                 Next i









'sfr cautare comenzi validate care solicita pr intrata
                                   
                                   
                                   
                                   
                                   
                
                
                End If
        
        End If
        
           Next x
        
        
        
        
        
        
        
        
        
        
        
        
        
          'sfr stergere doar dc nu e partiala
        End If
        
        
        'realizare factura de storno dc e ff
        
         If cdafacturata = 1 Then
        
        
                        rand_cda_fact_listaff = 1 + Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("lista facturi").Range("bv:bv"), Numarcomanda.Value, ThisWorkbook.Sheets("lista facturi").Range("B:b"))
                        
                        rand_Actual = ThisWorkbook.Sheets("lista facturi").Cells(1, 1) + 2
                        
                        MsgBox "rand de copiat=" & rand_cda_fact_listaff
                        MsgBox "rand actual=" & rand_Actual
                        
                        
                        For x = 2 To 80
                        
                        
                        If x >= 20 And x <= 28 Then
                        
                        ThisWorkbook.Sheets("lista facturi").Cells(rand_Actual, x) = 0 - ThisWorkbook.Sheets("lista facturi").Cells(rand_cda_fact_listaff, x)
                        
                       
                        
                        Else
                        ThisWorkbook.Sheets("lista facturi").Cells(rand_Actual, x) = ThisWorkbook.Sheets("lista facturi").Cells(rand_cda_fact_listaff, x)
                        
                         End If
                        
                        Next x
                        
                        
                        'precizare la ce cda e stornarea
                        ThisWorkbook.Sheets("lista facturi").Cells(rand_Actual, 81) = Numarcomanda.Value
                         'sfr precizare la ce cda e stornarea
        
        
          End If
        
        
        
        
        
        
        
        'sfr realizare factura de storno dc e ff



End If


Validarecomanda.Locked = True



End Sub

Private Sub nr1_Change()

End Sub



Private Sub nr3_Change()

End Sub

Function comandalocked()


                                marciauto.Locked = True
                                Modeleauto.Locked = True


                                     Numeclient.Locked = True
                                     Prenumeclient.Locked = True
                                     Telefonclient.Locked = True
                                     
                                     Adresaclient.Locked = True
                                     Buletinclient.Locked = True
                                    
                                     
                                     Marcamasina.Locked = True
                                     Modelmasina.Locked = True
                                     Masinainmatriculare.Locked = True
                                     Kilometraj.Locked = True
                                     
                                     nr1.Locked = True
                                     cod1.Locked = True
                                     denumire1.Locked = True
                                     cantitate1.Locked = True
                                     pret1.Locked = True
                                     total1.Locked = True
                                     piesesosite1.Locked = True
                                     
                                      nr2.Locked = True
                                     cod2.Locked = True
                                     denumire2.Locked = True
                                     cantitate2.Locked = True
                                     pret2.Locked = True
                                     total2.Locked = True
                                     piesesosite2.Locked = True
                                     
                                      nr3.Locked = True
                                     cod3.Locked = True
                                     denumire3.Locked = True
                                     cantitate3.Locked = True
                                     pret3.Locked = True
                                     total3.Locked = True
                                     piesesosite3.Locked = True
                                     
                                     
                                      nr4.Locked = True
                                     cod4.Locked = True
                                     denumire4.Locked = True
                                     cantitate4.Locked = True
                                     pret4.Locked = True
                                     total4.Locked = True
                                     piesesosite4.Locked = True
                                     
                                     
                                      nr5.Locked = True
                                     cod5.Locked = True
                                     denumire5.Locked = True
                                     cantitate5.Locked = True
                                     pret5.Locked = True
                                     total5.Locked = True
                                     piesesosite5.Locked = True
                                     
                                     
                                      nr6.Locked = True
                                     cod6.Locked = True
                                     denumire6.Locked = True
                                     cantitate6.Locked = True
                                     pret6.Locked = True
                                     total6.Locked = True
                                     piesesosite6.Locked = True
                                     
                                     
                                     nr7.Locked = True
                                     cod7.Locked = True
                                     denumire7.Locked = True
                                     cantitate7.Locked = True
                                     pret7.Locked = True
                                     total7.Locked = True
                                     piesesosite7.Locked = True
                                     
                                     
                                     nr8.Locked = True
                                     cod8.Locked = True
                                     denumire8.Locked = True
                                     cantitate8.Locked = True
                                     pret8.Locked = True
                                     total8.Locked = True
                                     piesesosite8.Locked = True
                                     
                                     nr9.Locked = True
                                     cod9.Locked = True
                                     denumire9.Locked = True
                                     cantitate9.Locked = True
                                     pret9.Locked = True
                                     total9.Locked = True
                                     piesesosite9.Locked = True


End Function

Function comandaunlocked()
                                
                                
                                    marciauto.Locked = False
                                    Modeleauto.Locked = False

                                     Numeclient.Locked = False
                                     Prenumeclient.Locked = False
                                     Telefonclient.Locked = False
                                     
                                     Adresaclient.Locked = False
                                     Buletinclient.Locked = False
                                    
                                     
                                     Marcamasina.Locked = False
                                     Modelmasina.Locked = False
                                     Masinainmatriculare.Locked = False
                                     Kilometraj.Locked = False
                                     
                                     nr1.Locked = False
                                     cod1.Locked = False
                                     denumire1.Locked = False
                                     cantitate1.Locked = False
                                     pret1.Locked = False
                                     total1.Locked = True
                                     piesesosite1.Locked = True
                                     
                                      nr2.Locked = False
                                     cod2.Locked = False
                                     denumire2.Locked = False
                                     cantitate2.Locked = False
                                     pret2.Locked = False
                                     total2.Locked = True
                                     piesesosite2.Locked = True
                                     
                                      nr3.Locked = False
                                     cod3.Locked = False
                                     denumire3.Locked = False
                                     cantitate3.Locked = False
                                     pret3.Locked = False
                                     total3.Locked = True
                                     piesesosite3.Locked = True
                                     
                                     
                                      nr4.Locked = False
                                     cod4.Locked = False
                                     denumire4.Locked = False
                                     cantitate4.Locked = False
                                     pret4.Locked = False
                                     total4.Locked = True
                                     piesesosite4.Locked = True
                                     
                                     
                                      nr5.Locked = False
                                     cod5.Locked = False
                                     denumire5.Locked = False
                                     cantitate5.Locked = False
                                     pret5.Locked = False
                                     total5.Locked = True
                                     piesesosite5.Locked = True
                                     
                                     
                                      nr6.Locked = False
                                     cod6.Locked = False
                                     denumire6.Locked = False
                                     cantitate6.Locked = False
                                     pret6.Locked = False
                                     total6.Locked = True
                                     piesesosite6.Locked = True
                                     
                                       nr7.Locked = False
                                     cod7.Locked = False
                                     denumire7.Locked = False
                                     cantitate7.Locked = False
                                     pret7.Locked = False
                                     total7.Locked = True
                                     piesesosite7.Locked = True
                                     
                                     
                                       nr8.Locked = False
                                     cod8.Locked = False
                                     denumire8.Locked = False
                                     cantitate8.Locked = False
                                     pret8.Locked = False
                                     total8.Locked = True
                                     piesesosite8.Locked = True
                                     
                                     
                                    nr9.Locked = False
                                     cod9.Locked = False
                                     denumire9.Locked = False
                                     cantitate9.Locked = False
                                     pret9.Locked = False
                                     total9.Locked = True
                                     piesesosite9.Locked = True


End Function

Function numarulcomenziidinsheet()

Dim nrcda As Integer

 nrcda = ThisWorkbook.Sheets("comenzisimplu").Cells(1, 1)

numarulcomenziidinsheet = nrcda

End Function

Private Sub Numarcomanda_Change()




            If Numarcomanda.Text = "" Or Numarcomanda.Text = "0" Or Not IsNumeric(Numarcomanda.Value) Then
            
            'MsgBox "la nr 1"
            Numarcomanda.Value = 1
            
            End If
            
            
             
              
           


                Dim nr_cda As Integer
                nr_cda = Numarcomanda.Text
                
                
                
                

                
        
                If nr_cda > numarulcomenziidinsheet + 1 Then
                
                
                
              ''  MsgBox "mai mare"
              ''  MsgBox "numar comanda=" & Numarcomanda.Value
              ''  MsgBox " max din tabel=" & numarulcomenziidinsheet
                
            
                Numarcomanda.Value = 1 + numarulcomenziidinsheet
                
                
                
               
      
        
                End If
                
                
                 
                 
                 x = Numarcomanda.Value + 1
                 
                 
                 'daca exista detalii se vor prelua in comanda
                 
                 
                 exista_cda = 1 + Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:b"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzisimplu").Range("B:b"))
                 
                 If exista_cda > 1 Then
                 
                 'comanda exista, trebuie preluate info dc exista
                                   ' MsgBox "exista"
                 
                 
                                     Numeclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 3)
                                     Prenumeclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 4)
                                     Telefonclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 5)
                                     
                                     Adresaclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 6)
                                     Buletinclient.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 7)
                                    
                                     
                                     Marcamasina.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 8)
                                     Modelmasina.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 9)
                                     Masinainmatriculare.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 11)
                                     Kilometraj.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 10)
                                     DDR.Text = ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 19)
                                     
                                     marciauto.Text = Marcamasina.Text
                                     Modeleauto.Text = Modelmasina.Text
                                     
                 
                 
                 
                 'dc nu exista introdusa se pune ddr si nr, user, statut partiala, modificare statut
                 Else
                 
                 
                 
                    ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 19) = CDate(Now())
                    
                    
                    DDR.Value = Now()
                    
                    an = Year(DDR.Value)
                    luna = Month(DDR.Value)
                    zi = Day(DDR.Value)
                    
                    ddrafisat = zi & "/" & luna & "/" & an
                    
                    DDR.Text = ddrafisat
                    
                    
                 
                    ThisWorkbook.Sheets("comenzisimplu").Cells(x, 2) = Numarcomanda.Value
                    ThisWorkbook.Sheets("comenzisimplu").Cells(x, 18) = User.Caption
                    ThisWorkbook.Sheets("comenzisimplu").Cells(x, 13) = 1
                    ThisWorkbook.Sheets("Comenzisimplu").Cells(x, 19) = DDR.Text
                    
                    poz_mod_statut = ThisWorkbook.Sheets("Modificare statut comenzi").Cells(1, 1) + 2
                    
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(poz_mod_statut, 2) = poz_mod_statut - 1
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(poz_mod_statut, 3) = Numarcomanda.Value
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(poz_mod_statut, 4) = "Comanda noua"
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(poz_mod_statut, 5) = CDate(Now())
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(poz_mod_statut, 6) = User.Caption
                    
                    
                                     Numeclient.Text = ""
                                     Prenumeclient.Text = ""
                                     Telefonclient.Text = ""
                                     
                                     Adresaclient.Text = ""
                                     Buletinclient.Text = ""
                                    
                                     
                                     Marcamasina.Text = ""
                                     Modelmasina.Text = ""
                                     Masinainmatriculare.Text = ""
                                     Kilometraj.Text = ""
                                     marciauto.Text = Marcamasina.Text
                                     Modeleauto.Text = Modelmasina.Text
                 
                 
                 
                 End If
                 'sfr dc nu exista introdusa se pune ddr si nr
                 
                 
                 'sfr daca exista detalii se vor prelua in comanda
        
        
                
                For j = 1 To 9
                
                                 If randdetaliicomanda(x - 1, j) > 1 Then
                               
                                 y = randdetaliicomanda(x - 1, j)
                               
                                 inutil = preluarepiese(y, 1)
                               
                               
                               
                            Else
                            
                            resetarepiese (j)
                            
                            
                            
                            
                            
                            
                            
                            End If
      
          
                 Next j
         
         
            

        

     If cdastearsa = 1 Then
                
                comandalocked
                
             '   MsgBox "cda stearsa"
                
                
                            stearsa.Value = True
                            comandapartiala.Value = False
                            asteptarepiese.Value = False
                            stearsa.Locked = True
                            
                            Livrata.Locked = True
                            Facturata.Locked = True
                            comandapartiala.Locked = True
                            Delivrat.Locked = True
                            
                            
                            Validarecomanda.Locked = True
                            
                         '   MsgBox "cda stearsa"
                            
                            
                            



                ElseIf cdavalidata = 1 And cdadelivrat = 1 And cdalivrata = 0 Then
                    
                                          '  MsgBox "delivrat"
                                            'MsgBox Numarcomanda.Value
                                            
                                            Delivrat.Value = True
                                            
                                            comandalocked
                                            
                                             asteptarepiese.Locked = True
                                           
                                            Facturata.Locked = True
                                            comandapartiala.Locked = True
                                            Livrata.Locked = True
                                            
                                            
                                            Validarecomanda.Locked = True
                                            
                                            
                                            
                                            
                                                            ElseIf cdadelivrat = 0 And cdavalidata = 1 Then
                                                            
                                                            comandalocked
                                                            
                                                            Livrata.Locked = True
                                                            Facturata.Locked = True
                                                            comandapartiala.Locked = True
                                                            Delivrat.Locked = True
                                                            
                                                            Validarecomanda.Locked = True
                                                            asteptarepiese.Value = True
                                            
                                            
                                                            ElseIf cdavalidata = 0 Then
                                                            
                                                            
                                                           ' MsgBox "partiala"
                                                            Validarecomanda.Locked = False
                                                            stearsa.Locked = True
                                                            Livrata.Locked = True
                                                            Facturata.Locked = True
                                                            Delivrat.Locked = True
                                                            comandapartiala.Value = True
                                                            
                                                            
                                                            comandaunlocked
                                                            
                                                            
                                                            
                                                   End If
                            
    





                    
          
                               
                             
                   




End Sub



Private Sub Numeclient_Change()

  ' ThisWorkbook.Sheets("selectie numeclienti").Cells(1, 12) = Numeclient.Text
    

   If cdavalidata = 0 Then
   
 '  MsgBox "partiala!"

   ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 3) = Numeclient.Text



        End If
        
        
   'selectare prenume clienti



    criteriulnume = Numeclient.Text


    ThisWorkbook.Sheets("selectie numeclienti filtre").UsedRange.AutoFilter _
    field:=1, _
    Criteria1:=criteriulnume, _
    VisibleDropDown:=False
 
      ThisWorkbook.Sheets("selectie prenumeclienti filtre").UsedRange.Clear
  
      ThisWorkbook.Sheets("selectie numeclienti filtre").UsedRange.Copy Destination:=ThisWorkbook.Sheets("selectie prenumeclienti filtre").Range("a1")
      
      


            If ThisWorkbook.Sheets("selectie numeclienti filtre").AutoFilterMode Then
             
                ThisWorkbook.Sheets("selectie numeclienti filtre").AutoFilterMode = False
             
            End If
            
            Prenumeclient.Text = ThisWorkbook.Sheets("selectie prenumeclienti filtre").Cells(2, 2)

'sfr selectare prenume clienti
        
   

      
        
        

End Sub

Private Sub piesesosite1_Change()


If Not IsNumeric(piesesosite1.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

piesesosite1.Text = ""

End If



'introducere detalii in comanda detaliata
                
     

'introducere detalii in comanda detaliata
Call detaliicomanda(6, 1)



upgradenrprsosite



End Sub
Private Sub piesesosite2_Change()


If Not IsNumeric(piesesosite2.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

piesesosite2.Text = ""

End If



    'introducere detalii in comanda detaliata
        Call detaliicomanda(6, 2)

upgradenrprsosite


End Sub
Private Sub piesesosite3_Change()


If Not IsNumeric(piesesosite3.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

piesesosite3.Text = ""

End If

        'introducere detalii in comanda detaliata
Call detaliicomanda(6, 3)

upgradenrprsosite

End Sub
Private Sub piesesosite4_Change()


If Not IsNumeric(piesesosite4.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

piesesosite4.Text = ""

End If
        
        
        'introducere detalii in comanda detaliata
Call detaliicomanda(6, 4)



upgradenrprsosite

End Sub
Private Sub piesesosite5_Change()


If Not IsNumeric(piesesosite5.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

piesesosite5.Text = ""

End If

        'introducere detalii in comanda detaliata
Call detaliicomanda(6, 5)

upgradenrprsosite

End Sub

Private Sub piesesosite7_Change()



Call detaliicomanda(6, 7)



End Sub

Private Sub piesesosite8_Change()



Call detaliicomanda(6, 8)



End Sub

Private Sub piesesosite9_Change()



Call detaliicomanda(6, 9)



End Sub
Private Sub piesesosite6_Change()


If Not IsNumeric(piesesosite6.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

piesesosite1.Text = ""

End If

    'introducere detalii in comanda detaliata
    Call detaliicomanda(6, 6)

upgradenrprsosite

End Sub




Private Sub Prenumeclient_Change()

If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 4) = Prenumeclient.Text


End If






End Sub

Private Sub TextBox1_Change()

End Sub



Private Sub pret1_Change()


                    If pret1.Text = "" Then
                    total1.Text = ""
                    
                    
                    
                    End If

            
            If Not IsNumeric(pret1.Value) Then
            
            'MsgBox "Pretul trebuie sa fie numeric!"
            
            pret1.Text = ""
            
            End If

                            
                            If IsNumeric(pret1.Value) And IsNumeric(cantitate1.Value) Then
                            
                            total1.Text = pret1 * cantitate1
                            
                            
                            End If
                            
                            
                         Call detaliicomanda(4, 1)
                            
                            



End Sub

Private Sub pret2_Change()


     If pret2.Text = "" Then
                    total2.Text = ""
                    
                    
                    
                    End If


If Not IsNumeric(pret2.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret2.Text = ""

End If

If IsNumeric(pret2.Value) And IsNumeric(cantitate2.Value) Then

total2.Text = pret2 * cantitate2


End If


            Call detaliicomanda(4, 2)



End Sub

Private Sub pret3_Change()


                    If pret3.Text = "" Then
                    total3.Text = ""
                    
                    
                    
                    End If


If Not IsNumeric(pret3.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret3.Text = ""

End If


If IsNumeric(pret3.Value) And IsNumeric(cantitate3.Value) Then

total3.Text = pret3 * cantitate3


End If



         Call detaliicomanda(4, 3)



End Sub

Private Sub pret4_Change()




                    If pret4.Text = "" Then
                    total4.Text = ""
                    
                    
                    
                    End If


If Not IsNumeric(pret4.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret4.Text = ""

End If


If IsNumeric(pret4.Value) And IsNumeric(cantitate4.Value) Then

total4.Text = pret4 * cantitate4


End If


        Call detaliicomanda(4, 4)


End Sub

Private Sub pret5_Change()


    

                    If pret5.Text = "" Then
                    total5.Text = ""
                    
                    
                    
                    End If


If Not IsNumeric(pret5.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret5.Text = ""

End If


If IsNumeric(pret5.Value) And IsNumeric(cantitate5.Value) Then

total5.Text = pret5 * cantitate5


End If


    Call detaliicomanda(4, 5)


End Sub
Private Sub pret6_Change()


    

                    If pret6.Text = "" Then
                    total6.Text = ""
                    
                    
                    
                    End If
                    


If Not IsNumeric(pret6.Value) Then

'MsgBox "Pretul trebuie sa fie numeric!"

pret6.Text = ""

End If

If IsNumeric(pret6.Value) And IsNumeric(cantitate6.Value) Then

total6.Text = pret6 * cantitate6


End If

    'introducere detalii in comanda detaliata
    Call detaliicomanda(4, 6)



End Sub

Private Sub Rulaj_Click()

End Sub

Private Sub statutcomanda_Click()

End Sub



Private Sub Telefon_Click()

End Sub

Private Sub Telefonclient_Change()


If cdavalidata <> 1 Then

ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 5) = Telefonclient.Text


End If



End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox12_Change()

End Sub

Private Sub total1_Change()

'introducere detalii in comanda detaliata
 Call detaliicomanda(5, 1)


actualizaretotalpr (1)
        


End Sub

Private Sub total2_Change()

'introducere detalii in comanda detaliata
Call detaliicomanda(5, 2)
actualizaretotalpr (2)


End Sub
Private Sub total3_Change()

'introducere detalii in comanda detaliata
Call detaliicomanda(5, 3)
actualizaretotalpr (3)


End Sub
Private Sub total4_Change()

'introducere detalii in comanda detaliata
Call detaliicomanda(5, 4)
actualizaretotalpr (4)

End Sub
Private Sub total5_Change()

'introducere detalii in comanda detaliata
Call detaliicomanda(5, 5)
actualizaretotalpr (5)


End Sub
Private Sub total6_Change()

'introducere detalii in comanda detaliata
Call detaliicomanda(5, 6)
actualizaretotalpr (6)

End Sub

Private Sub total7_Change()


Call detaliicomanda(5, 7)

Call actualizaretotalmo




End Sub

Private Sub total8_Change()


Call detaliicomanda(5, 8)
Call actualizaretotalmo

End Sub


Private Sub total9_Change()


Call detaliicomanda(5, 9)
Call actualizaretotalmo

End Sub

Private Sub totalcda_Change()



            valoare_tva = ThisWorkbook.Sheets("intrare piese").Cells(1, 15)
            
            totalcdatva.Value = totalcda.Value * valoare_tva / 100
            totalcdatva.Text = totalcdatva.Value
            
            
          
          totalplatatexta = 0
            
            
         If IsNumeric(totalcda.Value) And IsNumeric(totalcdatva.Value) Then
        
            totalplatatexta = totalcda.Value * 1 + totalcdatva.Value * 1
        
        
        
         End If
        
        
        

            totalplata.Value = totalplatatexta
            totalplata.Text = totalplatatexta




End Sub

Private Sub totalcdatva_Change()

valoare_tva = ThisWorkbook.Sheets("intrare piese").Cells(1, 15)
            
            totalcdatva.Value = totalcda.Value * valoare_tva / 100
            totalcdatva.Text = totalcdatva.Value
            
            
        
            
            
        


            



End Sub

Private Sub totalmo_Change()

            valoare_tva = ThisWorkbook.Sheets("intrare piese").Cells(1, 15)
            
            totaltvamo.Value = totalmo.Value * valoare_tva / 100
            totaltvamo.Text = totaltvamo.Value
            
            
         totalcdatext = 0

        
        
        If IsNumeric(totalpr.Value) Then
        
        totalcdatext = totalpr.Value + totalcdatext
        
        
        
        End If
        
          If IsNumeric(totalmo.Value) Then
        
        totalcdatext = totalmo.Value + totalcdatext
        
        
        
        End If
        
        
        

            totalcda.Value = totalcdatext
            totalcda.Text = totalcdatext

            

    

          


End Sub

Private Sub totalplata_Change()

End Sub

Private Sub totalpr_Change()



valoare_tva = ThisWorkbook.Sheets("intrare piese").Cells(1, 15)

            
            totaltvapr.Value = totalpr.Value * valoare_tva / 100
            totaltvapr.Text = totaltvapr.Value

        totalcdatext = 0

        
        
        If IsNumeric(totalpr.Value) Then
        
        totalcdatext = totalpr.Value + totalcdatext
        
        
        
        End If
        
          If IsNumeric(totalmo.Value) Then
        
        totalcdatext = totalmo.Value + totalcdatext
        
        
        
        End If
        
        
        

            totalcda.Value = totalcdatext
            totalcda.Text = totalcdatext

End Sub
Function actualizaretotalpr(poz)
    
        actualizaretotalpr = 1
        
        totalprtext = 0
        
        If IsNumeric(total1.Value) Then
        
        totalprtext = total1.Value + totalprtext
        
        
        
        End If
        
        
         If IsNumeric(total2.Value) Then
        
        totalprtext = total2.Value + totalprtext
        
        
        
        End If
        
         If IsNumeric(total3.Value) Then
        
        totalprtext = total3.Value + totalprtext
        
        
        
        End If
        
         If IsNumeric(total4.Value) Then
        
        totalprtext = total4.Value + totalprtext
        
        
        
        End If
        
         If IsNumeric(total5.Value) Then
        
        totalprtext = total5.Value + totalprtext
        
        
        
        End If
        
         If IsNumeric(total6.Value) Then
        
        totalprtext = total6.Value + totalprtext
        
        
        
        End If
        
        
        
        
        
        
        
        totalpr.Text = totalprtext
        





End Function
Function actualizaretotalmo()
    
        
        
        totalmotext = 0
        
        If IsNumeric(total7.Value) Then
        
        totalmotext = total7.Value + totalmotext
        
        
        
        End If
        
        
         If IsNumeric(total8.Value) Then
        
        totalmotext = total8.Value + totalmotext
        
        
        
        End If
        
         If IsNumeric(total9.Value) Then
        
        totalmotext = total9.Value + totalmotext
        
        
        
        End If
        
        
        totalmo.Text = totalmotext
        
        actualizaretotalmo = 1




End Function

Private Sub totaltvamo_Change()

End Sub

Private Sub totaltvapr_Change()

End Sub

Private Sub UserForm_Click()


End Sub

Private Sub Validarecomanda_Click()








If Marcamasina.Text <> "" And Modelmasina.Text <> "" And Kilometraj.Text <> "" And Numarcomanda.Text <> "" And Numeclient <> "" And Numeclient <> "" And Telefonclient <> "" And Adresaclient <> "" And Buletinclient <> "" Then


If Not IsNumeric(Kilometraj.Value) Then


               MsgBox "Kilometrajul trebuie sa fie numeric"

    Else
    
        
    
    
       'se verifica daca randurile sunt completate
       
        totalranduri = 0
        
        
        eroare1 = 0
        
        rand1 = 0
        
        If cod1.Text <> "" Then
        rand1 = rand1 + 1
        End If
        
        
        If IsNumeric(cantitate1.Text) And cantitate1.Text <> "" Then
        rand1 = rand1 + 1
        End If
        
        
         If denumire1.Text <> "" Then
        rand1 = rand1 + 1
        End If
        
        
          
        If IsNumeric(pret1.Text) And pret1.Text <> "" Then
        rand1 = rand1 + 1
        End If
        
           If IsNumeric(total1.Text) And total1.Text <> "" Then
        rand1 = rand1 + 1
        End If
        
        
        
        If rand1 > 0 And rand1 < 5 Then
        
        eroare1 = 1
        eroaretotalavalidare = 1
        
        
        
        'msgbox rand1
        
        MsgBox "Randul 1 de piese nu este completat"
        
        End If
        
        
        
            If rand1 = 5 Then
            totalranduri = totalranduri + 1
            End If
            
            
            
            
            eroare2 = 0
        
        rand2 = 0
        
        If cod2.Text <> "" Then
        rand2 = rand2 + 1
        End If
        
        
        If IsNumeric(cantitate2.Text) And cantitate2.Text <> "" Then
        rand2 = rand2 + 1
        End If
        
        
         If denumire2.Text <> "" Then
        rand2 = rand2 + 1
        End If
        
        
          
        If IsNumeric(pret2.Text) And pret2.Text <> "" Then
        rand2 = rand2 + 1
        End If
        
           If IsNumeric(total2.Text) And total2.Text <> "" Then
        rand2 = rand2 + 1
        End If
        
        
        
        If rand2 > 0 And rand2 < 5 Then
        
        eroare2 = 1
        eroaretotalavalidare = 1
        
        
        
        'msgbox rand2
        
        MsgBox "Randul 2 de piese nu este completat"
        
        End If
        
        
        
            If rand2 = 5 Then
            totalranduri = totalranduri + 1
            End If
            
            
            
            
            eroare3 = 0
        
        rand3 = 0
        
        If cod3.Text <> "" Then
        rand3 = rand3 + 1
        End If
        
        
        If IsNumeric(cantitate3.Text) And cantitate3.Text <> "" Then
        rand3 = rand3 + 1
        End If
        
        
         If denumire3.Text <> "" Then
        rand3 = rand3 + 1
        End If
        
        
          
        If IsNumeric(pret3.Text) And pret3.Text <> "" Then
        rand3 = rand3 + 1
        End If
        
           If IsNumeric(total3.Text) And total3.Text <> "" Then
        rand3 = rand3 + 1
        End If
        
       
        
        If rand3 > 0 And rand3 < 5 Then
        
        eroare3 = 1
        eroaretotalavalidare = 1
        
        
        
        'msgbox rand3
        
        MsgBox "Randul 3 de piese nu este completat"
        
        End If
        
        
        
            If rand3 = 5 Then
            totalranduri = totalranduri + 1
            End If
            
            
            
            
            eroare4 = 0
        
        rand4 = 0
        
        If cod4.Text <> "" Then
        rand4 = rand4 + 1
        End If
        
        
        If IsNumeric(cantitate4.Text) And cantitate4.Text <> "" Then
        rand4 = rand4 + 1
        End If
        
        
         If denumire4.Text <> "" Then
        rand4 = rand4 + 1
        End If
        
        
          
        If IsNumeric(pret4.Text) And pret4.Text <> "" Then
        rand4 = rand4 + 1
        End If
        
           If IsNumeric(total4.Text) And total4.Text <> "" Then
        rand4 = rand4 + 1
        End If
        
        
        
        If rand4 > 0 And rand4 < 5 Then
        
        eroare4 = 1
        eroaretotalavalidare = 1
        
        
        
        'msgbox rand4
        
        MsgBox "Randul 4 de piese nu este completat"
        
        End If
        
        
        
            If rand4 = 5 Then
            totalranduri = totalranduri + 1
            End If
            
            
            
            eroare5 = 0
        
        rand5 = 0
        
        If cod5.Text <> "" Then
        rand5 = rand5 + 1
        End If
        
        
        If IsNumeric(cantitate5.Text) And cantitate5.Text <> "" Then
        rand5 = rand5 + 1
        End If
        
        
         If denumire5.Text <> "" Then
        rand5 = rand5 + 1
        End If
        
        
          
        If IsNumeric(pret5.Text) And pret5.Text <> "" Then
        rand5 = rand5 + 1
        End If
        
           If IsNumeric(total5.Text) And total5.Text <> "" Then
        rand5 = rand5 + 1
        End If
        
        
        
        If rand5 > 0 And rand5 < 5 Then
        
        eroare5 = 1
        eroaretotalavalidare = 1
        
        
        
        'msgbox rand5
        
        MsgBox "Randul 5 de piese nu este completat"
        
        End If
        
        
        
            If rand5 = 6 Then
            totalranduri = totalranduri + 1
            End If
            
            
            eroare6 = 0
        
        rand6 = 0
        
        If cod6.Text <> "" Then
        rand6 = rand6 + 1
        End If
        
        
        If IsNumeric(cantitate6.Text) And cantitate6.Text <> "" Then
        rand6 = rand6 + 1
        End If
        
        
         If denumire6.Text <> "" Then
        rand6 = rand6 + 1
        End If
        
        
          
        If IsNumeric(pret6.Text) And pret6.Text <> "" Then
        rand6 = rand6 + 1
        End If
        
           If IsNumeric(total6.Text) And total6.Text <> "" Then
        rand6 = rand6 + 1
        End If
        
        
        
        If rand6 > 0 And rand6 < 5 Then
        
        eroare6 = 1
        eroaretotalavalidare = 1
        
        
        
        'msgbox rand6
        
        MsgBox "Randul 6 de piese nu este completat"
        
        End If
        
        
        
            If rand6 = 5 Then
            totalranduri = totalranduri + 1
            End If
            
            
            
            
            'se verifica daca este vreo piesa ceruta si daca exista erori de completare incompleta a liniilor de piese
            
            If totalranduri = 0 Then
            MsgBox "Nu este ceruta nici o piesa"
            End If
            
    
            If eroaretotalavalidare = 0 And totalranduri > 0 And cdavalidata = 0 Then
    
    
              MsgBox "Comanda a fost validata!"
              
              
              
              
              
              
              
              comandalocked
              
              
              comandapartiala.Locked = True
              Delivrat.Locked = True
              
              
              
              
              'actualizare statut comanda
                    
                    randcomandamodificata = ThisWorkbook.Sheets("Modificare statut comenzi").Cells(1, 1) + 2
                    
     
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(randcomandamodificata, 2) = randcomandamodificata - 1
                    
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(randcomandamodificata, 3) = Numarcomanda.Value
                    
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(randcomandamodificata, 4) = "Comanda validata"
                    
                     ThisWorkbook.Sheets("Modificare statut comenzi").Cells(randcomandamodificata, 5) = Now()
                     
                    ThisWorkbook.Sheets("Modificare statut comenzi").Cells(randcomandamodificata, 6) = User.Caption
              
              
              asteptarepiese.Value = True
              
              
              
               Dim datavalidata As Date
            
               datavalidata = True
            
                  
                Dim datavalidata2 As Date
                datavalidata2 = Format(Now(), "dd-MMMM, yyyy")
        
        
        
        
                    DDR.Value = CDate(datavalidata2)
        
        
                    an = Year(DDR.Value)
                    luna = Month(DDR.Value)
                    zi = Day(DDR.Value)
                    
                    ddrafisat = zi & "/" & luna & "/" & an
                    
                    DDR.Text = ddrafisat
        
        
        
                 Validarecomanda.Locked = True

                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 2) = UCase(Numarcomanda.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 3) = UCase(Numeclient.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 4) = UCase(Prenumeclient.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 5) = UCase(Telefonclient.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 6) = UCase(Adresaclient.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 7) = UCase(Buletinclient.Text)
                 ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 8) = UCase(Marcamasina.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 9) = UCase(Modelmasina.Text)
                       
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 10) = UCase(Kilometraj.Text)
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 11) = UCase(Masinainmatriculare.Text)
                 ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 20) = UCase(DDR.Text)
                

      
        
        


          If ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 25) <> 1 Then
    
          ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Text + 1, 14) = 1
        
           End If
        
   
             
             
             
             
             'actualizare valoare p_sosite
             
              For x = 1 To 6
              
              
              qnt_iesire_stoc = 0
              
              'msgbox "aici pr sosite"
        
                'msgbox "pr sosite=" & Me.Controls("piesesosite" & x).Value
        
                If Me.Controls("cantitate" & x).Text <> "" And Me.Controls("cantitate" & x).Value <> 0 Then
                
                
                                    pr_intrare = Me.Controls("cod" & x).Text
                                    qnt_intrare = Me.Controls("cantitate" & x).Value
                                    denumire_pr = Me.Controls("denumire" & x).Text
                                    
                                    
                                    restsosire = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("comenzi detaliate").Range("r:r"), ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzi detaliate").Range("e:e"), pr_intrare)
                                    
                                    MsgBox "reper=" & pr_intrare & "  rest sosire=" & restsosire
                                    
                                    
                                    
                                    
                                    
                                     Set cautare_pr_stoc = ThisWorkbook.Sheets("stoc piese").Range("c1")
                                    Set cautare_pr_stoc = ThisWorkbook.Sheets("stoc piese").Columns(3).Find(What:=pr_intrare, After:=cautare_pr_stoc, _
                                    LookIn:=xlValues, LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                         MatchCase:=True)
                                    
                                        
                                     existapiesa = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("stoc piese").Columns(3), Me.Controls("cod" & x).Text, ThisWorkbook.Sheets("stoc piese").Columns(2))
                                        
                                        If existapiesa > 0 Then
                                        
                                        rand_pr_stoc = cautare_pr_stoc.Row
                                        
                                        qnt_iesire_stoc = Application.WorksheetFunction.Min(ThisWorkbook.Sheets("stoc piese").Cells(rand_pr_stoc, 4), restsosire)
                                        MsgBox "cantitate iesire stoc=" & qnt_iesire_stoc
                                         
                                        
                                        
                                        ThisWorkbook.Sheets("stoc piese").Cells(rand_pr_stoc, 4) = ThisWorkbook.Sheets("stoc piese").Cells(rand_pr_stoc, 4) - qnt_iesire_stoc
                                        
                                        If Me.Controls("piesesosite" & x).Text = "" Then
                                        Me.Controls("piesesosite" & x).Value = 0
                                        End If
                                        
                                        
                                        Me.Controls("piesesosite" & x).Value = Me.Controls("piesesosite" & x).Value + qnt_iesire_stoc
                                        
                                        
                                        
                                        End If
             
             End If
             
       
             
             
             'sfr actualizare valoare sosite
             
             
              'actualizare c_det
                                    
        

                                            Dim iLoop2 As Integer
        

                                            Dim cautare_dt_det As Range
                                            Set cautare_dt_det = ThisWorkbook.Sheets("comenzi detaliate").Range("c1")
                                           
                
                
                                         iLoop2 = WorksheetFunction.CountIf(ThisWorkbook.Sheets("comenzi detaliate").Columns(3), Numarcomanda.Value)
                                         
                                         MsgBox "iloop2=" & iLoop2
                                         
                                         For j = 1 To iLoop2
                                    
                                       
                                        Set cautare_dt_det = ThisWorkbook.Sheets("comenzi detaliate").Columns(3).Find(What:=Numarcomanda.Value, After:=cautare_dt_det, _
                                        LookIn:=xlValues, LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                         MatchCase:=True)
                                    
                                        
                                        
                                        
                                        
                                        rand_dt_det = cautare_dt_det.Row
                                        
                                        MsgBox "rand cda det=" & rand_dt_det
                                        
                                      
                                        
                                        If ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 5) = pr_intrare And ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 10) < ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 6) Then
                                        
                                        
                                        MsgBox "cantitate iesire stoc=" & qnt_iesire_stoc
                                        
                                        
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 10) = ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 10) + qnt_iesire_stoc
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 18) = ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 6) - ThisWorkbook.Sheets("comenzi detaliate").Cells(rand_dt_det, 10)
                                        
                                        
                                        End If
                                    
                                  
                                          Next j
                                    
                                    
                                 
                                   'sfr actualizare c_det
                                   
                                   
                                   
                                    'actualizare iesire piese
                                    
                                    If qnt_iesire_stoc > 0 Then
                                         
                                         linieiesirepiese = ThisWorkbook.Sheets("iesire piese").Cells(1, 1) + 2
                                        
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 2) = ThisWorkbook.Sheets("iesire piese").Cells(1, 1) + 1
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 3) = pr_intrare
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 4) = qnt_iesire_stoc
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 5) = denumire_pr
                                         
                                         'nume client
                                         
                                         poznumeclient = 1 + Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzisimplu").Range("b:b"))
                                         
                                         
                                         
                                         'msgbox "poz nume client=" & poznumeclient
                                         'msgbox "nr comanda=" & nrcdaprgasita
                                         
                                         
                                          
                                         
                                         
                                         'msgbox "range client row=" & rangeclient.Row
                                         
                                         
                                         
                                         
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 6) = ThisWorkbook.Sheets("comenzisimplu").Cells(poznumeclient, 3)
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 7) = ThisWorkbook.Sheets("comenzisimplu").Cells(poznumeclient, 4)
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 8) = Me.Controls("pret" & x).Value
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 10) = Numarcomanda.Value
                                            
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 11) = CDate(Now())
                                            
                                            
                                            
                                        End If
                                    'sfr actualizare iesire piese
                                    
                                   
                               Next x
                                   
                                    'actualizare statut comanda de livrat
        
        nr_pr_cerute_cda = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzi detaliate").Range("f:f"))
        nr_pr_sosite_cda = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzi detaliate").Range("j:j"))
                
                
                
                If nr_pr_cerute_cda = nr_pr_sosite_cda Then
                
                MsgBox "comanda " & nrcdaprgasita & " este de livrat!"
                
                Delivrat.Value = True
                
                
                ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 15) = 1
                
                End If
                
                
               
             
    

'final de validare

End If

End If

Else: MsgBox "Toate campurile trebuie completate!"



End If




End Sub


Sub detaliicomanda(campul, randul)


                Select Case campul
                
                Case 1
                coloanaexcel = 5
                rez = "cod"
                
                Case 2
                coloanaexcel = 6
                rez = "cantitate"
                
                Case 3
                coloanaexcel = 7
                rez = "denumire"
                
                Case 4
                coloanaexcel = 8
                rez = "pret"
                
                Case 5
                coloanaexcel = 9
                rez = "total"
                
                Case 6
                coloanaexcel = 10
                rez = "piesesosite"
                
                End Select
                
                 '' MsgBox "rez=" & rez
               '' MsgBox "coloana=" & coloanaexcel


 'introducere detalii in comanda detaliata
                
            If ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 14) <> 1 Then
            
            randcomandadetaliata = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("comenzi detaliate").Range("B:B"), ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzi detaliate").Range("d:d"), randul)
            
            'MsgBox "rand comanda detaliata " & randcomandadetaliata
            
                        
            
                        If randcomandadetaliata = 0 Then
                        
                       ' MsgBox "randdetaliata=" & randcomandadetaliata
                       
                       
                        
                        randdetaliata = ThisWorkbook.Sheets("comenzi detaliate").Cells(1, 1) + 2
                        
                        ' MsgBox "randdetaliata=" & randcomandadetaliata
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 3) = Numarcomanda.Value
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 2) = ThisWorkbook.Sheets("comenzi detaliate").Cells(1, 1) + 1
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 4) = randul
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, coloanaexcel) = Me.Controls(rez & randul).Text
                        
                         ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 12) = CDate(Now())
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 13) = User.Caption
                        
                        If rez = "cantitate" Then
                        
                      '  MsgBox "se upgradeaza cantitatea"
                        
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 18) = Me.Controls(rez & randul).Value
                        
                        End If
                        
                        
                        
                        
                        
                        If randul < 7 Then
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 15) = 1
                        
                        
                       Else: ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 16) = 1
                        
                        End If
                        
                        
                        
                      
                        
                        
                           
                         ElseIf randcomandadetaliata > 0 Then
                          
                         '   MsgBox "randdetaliata=" & randcomandadetaliata
                          
                        '  MsgBox "aici"
                           
                           randdetaliata = 1 + Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("comenzi detaliate").Range("B:B"), ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), Numarcomanda.Value, ThisWorkbook.Sheets("comenzi detaliate").Range("d:d"), randul)
                           
                         '  MsgBox "randdetaliata=" & randdetaliata
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 3) = Numarcomanda.Value
                        
                      ' ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 2) = ThisWorkbook.Sheets("comenzi detaliate").Cells(1, 1) + 1
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 4) = randul
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, coloanaexcel) = Me.Controls(rez & randul).Text
                           
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 12) = CDate(Now())
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 13) = User.Caption
                        
                        If rez = "cantitate" Then
                        
                       ' MsgBox "se upgradeaza cantitatea"
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 18) = Me.Controls(rez & randul).Value
                        
                        End If
                        
                        
                         If randul < 7 Then
                        
                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 15) = 1
                        
                        
                       Else: ThisWorkbook.Sheets("comenzi detaliate").Cells(randdetaliata, 16) = 1
                        
                        End If
                        
                        
                           
                           End If
            
            
            
            
            
            End If
                
                'sfarsit introducere detalii in comanda detaliata
                
                
                
End Sub



Function randdetaliicomanda(cda, linia)


randdetaliicomanda = 1 + Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("comenzi detaliate").Range("B:b"), ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), cda, ThisWorkbook.Sheets("comenzi detaliate").Range("d:d"), linia)

                
                
                
                
End Function


Function preluarepiese(z, linie)

                            
    
    
    


                               Me.Controls("cod" & linie).Text = ThisWorkbook.Sheets("comenzi detaliate").Cells(z, 5)
                               Me.Controls("cantitate" & linie).Text = ThisWorkbook.Sheets("comenzi detaliate").Cells(z, 6)
                               Me.Controls("denumire" & linie).Text = ThisWorkbook.Sheets("comenzi detaliate").Cells(z, 7)
                               Me.Controls("pret" & linie).Text = ThisWorkbook.Sheets("comenzi detaliate").Cells(z, 8)
                               Me.Controls("total" & linie).Text = ThisWorkbook.Sheets("comenzi detaliate").Cells(z, 9)
                               Me.Controls("piesesosite" & linie).Text = ThisWorkbook.Sheets("comenzi detaliate").Cells(z, 10)



preluarepiese = 1

End Function


Function resetarepiese(linie)

                                Me.Controls("cod" & linie).Text = ""
                               Me.Controls("cantitate" & linie).Text = ""
                               Me.Controls("denumire" & linie).Text = ""
                               Me.Controls("pret" & linie).Text = ""
                               Me.Controls("total" & linie).Text = ""
                               Me.Controls("piesesosite" & linie).Text = ""









End Function
                
Function upgradenrprcerute()



            upgradenrprcerute = 0
            
            
            Dim totalnrcerute1 As Integer
            Dim totalnrcerute2 As Integer
            Dim totalnrcerute3 As Integer
            Dim totalnrcerute4 As Integer
            Dim totalnrcerute5 As Integer
            Dim totalnrcerute6 As Integer
            
            If IsNumeric(cantitate1.Value) Then
           totalnrcerute1 = cantitate1.Value
           
           End If
           
           
             If IsNumeric(cantitate2.Value) Then
           totalnrcerute2 = cantitate2.Value
           
           End If
           
           
             If IsNumeric(cantitate3.Value) Then
           totalnrcerute3 = cantitate3.Value
           
           End If
           
           
             If IsNumeric(cantitate4.Value) Then
           totalnrcerute4 = cantitate4.Value
           
           End If
           
           
             If IsNumeric(cantitate5.Value) Then
           totalnrcerute5 = cantitate5.Value
           
           End If
           
           
             If IsNumeric(cantitate6.Value) Then
           totalnrcerute6 = cantitate6.Value
           
           End If
           
           
          
           
           
           upgradenrprcerute = totalnrcerute1 + totalnrcerute2 + totalnrcerute3 + totalnrcerute4 + totalnrcerute5 + totalnrcerute6
             
              'MsgBox upgradenrprcerute



End Function



Function upgradenrprsosite()



            upgradenrprsosite = 0
            
            
            Dim totalnrsosite1 As Integer
            Dim totalnrsosite2 As Integer
            Dim totalnrsosite3 As Integer
            Dim totalnrsosite4 As Integer
            Dim totalnrsosite5 As Integer
            Dim totalnrsosite6 As Integer
            
            If IsNumeric(piesesosite1.Value) Then
           totalnrsosite1 = piesesosite1.Value
           
           End If
           
           
             If IsNumeric(piesesosite2.Value) Then
           totalnrsosite2 = piesesosite2.Value
           
           End If
           
           
             If IsNumeric(piesesosite3.Value) Then
           totalnrsosite3 = piesesosite3.Value
           
           End If
           
           
             If IsNumeric(piesesosite4.Value) Then
           totalnrsosite4 = piesesosite4.Value
           
           End If
           
           
             If IsNumeric(piesesosite5.Value) Then
           totalnrsosite5 = piesesosite5.Value
           
           End If
           
           
             If IsNumeric(piesesosite6.Value) Then
           totalnrsosite6 = piesesosite6.Value
           
           End If
           
           
          
           
           
           upgradenrprsosite = totalnrsosite1 + totalnrsosite2 + totalnrsosite3 + totalnrsosite4 + totalnrsosite5 + totalnrsosite6
           
           
           ThisWorkbook.Sheets("comenzisimplu").Cells(Numarcomanda.Value + 1, 24) = upgradenrprsosite
             
             
              'MsgBox upgradenrprsosite



End Function

Private Sub adaospiese_Change()



If Not IsNumeric(adaospiese.Text) Or adaospiese.Text = "0" Then
adaospiese.Text = ""

End If



If adaospiese.Text = "" Or adaospiese.Value <= 0 Then
adaospiese.Value = 40

End If




Sheets("Intrare piese").Cells(1, 12) = adaospiese.Text







End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub adaugareinstoc_Click()





If intrarecodpiesa.Text <> "" And denumireintrare.Text <> "" And cantitateintrare.Text <> "" And furnizorintrarepiese.Text <> "" And _
pretachizitieintrare.Text <> "" And pretvanzareintrare.Text <> "" And adaospiese.Text <> "" Then


                If IsNumeric(cantitateintrare.Text) And IsNumeric(pretachizitieintrare.Text) And IsNumeric(pretvanzareintrare.Text) And IsNumeric(adaospiese.Text) Then
                
                
                 MsgBox "Adaugata in stoc!"
                 
                 
                 'introducere detalii in intrare piese
                 
                 linieadaugare = Sheets("intrare piese").Cells(1, 1) + 2
                 
                 Dim datapiese As Date
                 
                 Dim pretvi As Double
                 pretvi = pretvanzareintrare.Value
                 
                 
                 MsgBox "linie=" & linieadaugare
                 
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 2) = ThisWorkbook.Sheets("intrare piese").Cells(1, 1) + 1
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 3) = intrarecodpiesa.Text
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 4) = denumireintrare.Text
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 5) = cantitateintrare.Value
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 6) = furnizorintrarepiese.Text
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 7) = pretachizitieintrare.Value
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 8) = pretvi
                  ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 9) = ThisWorkbook.Sheets("intrare piese").Cells(1, 15) / 100
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 11) = CDate(Now())
                 ThisWorkbook.Sheets("intrare piese").Cells(linieadaugare, 17) = ThisWorkbook.Sheets("start aplicatie").Cells(4, 2)
                 
                  'sfr introducere detalii in intrare piese
                  
                  'updatare stoc piese
                  
                  
                  
                  varstoc = Application.WorksheetFunction.SumIf(Sheets("stoc piese").Range("C:C"), intrarecodpiesa.Text, Sheets("stoc piese").Range("b:b"))
                  
                  MsgBox "stoc=" & varstoc
                  
                  'daca piesa nu a fost introdusa in stoc vreodata
                  
                  If varstoc = 0 Then
                  
                  liniestoc = Sheets("stoc piese").Cells(1, 1) + 2
                  Sheets("stoc piese").Cells(liniestoc, 2) = liniestoc - 1
                  Sheets("stoc piese").Cells(liniestoc, 3) = intrarecodpiesa
                  Sheets("stoc piese").Cells(liniestoc, 4) = cantitateintrare.Value
                  
                  
                  
                  
                  'daca piesa exista in stoc
                  
                  Else
                  
                  liniestoc = varstoc + 1
                  
                  Sheets("stoc piese").Cells(liniestoc, 3) = intrarecodpiesa
                  Sheets("stoc piese").Cells(liniestoc, 4) = cantitateintrare.Value + Sheets("stoc piese").Cells(liniestoc, 4)
                  
                  
                  
                  
                  
                  
                  End If
                  
                  
                  
                  
                  
                 'sfr updatare stoc piese
                 
                 
                 
                 
                 
                 End If
 
End If



'cautare comenzi validate care solicita pr intrata

'cautare comenzi validate cu piesa introdusa si cu nr sosit < nr cerut


        
        
        
        Dim nrcdagasita As Integer
        
        valstoc = ThisWorkbook.Sheets("stoc piese").Cells(liniestoc, 4)
        
        MsgBox "valstoc=" & valstoc
        
        referintapr = intrarecodpiesa.Text
        
        MsgBox "referinta pr=" & referintapr
        
                Dim iLoop As Integer
        

                Dim rNa As Range
                
                Dim i As Integer
                
                
                
                iLoop = WorksheetFunction.CountIf(ThisWorkbook.Sheets("comenzi detaliate").Columns(5), referintapr)
                
                Set rNa = ThisWorkbook.Sheets("comenzi detaliate").Range("E1")
                
                MsgBox "numar de comenzi gasite=" & iLoop
                
                'doar daca s-au gasit comenzi
                
                
                If iLoop > 0 Then
                
                 For i = 1 To iLoop
                 
                 
                
                  Set rNa = ThisWorkbook.Sheets("comenzi detaliate").Columns(5).Find(What:=referintapr, After:=rNa, _
                    LookIn:=xlValues, LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                 MatchCase:=True)
                
                      rNa.Font.Bold = True
                      
                      
                      
                      randulprgasita = rNa.Row
                      
                  '    MsgBox nrcdaprgasita & " " & randulprgasita & "  i=" & i
                      
                      
                      
                      
                      
                      
                      nrcdaprgasita = ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 3)
                      
                     MsgBox "nr comanda=" & nrcdaprgasita
                      
                      
                      statutcdaprgasita = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), nrcdaprgasita, ThisWorkbook.Sheets("comenzisimplu").Range("n:n"))
                      statutstearsa = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), nrcdaprgasita, ThisWorkbook.Sheets("comenzisimplu").Range("y:y"))
                      rest_de_sosit = Application.WorksheetFunction.SumIfs(ThisWorkbook.Sheets("comenzi detaliate").Range("r:r"), ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), nrcdaprgasita, ThisWorkbook.Sheets("comenzi detaliate").Range("e:e"), referintapr)
                      
                      
                      
                     ' MsgBox nrcdaprgasita & " " & statutcdaprgasita
                      
                      'actualizare stoc, rest de sosit, iesiri daca dt este validat
                      
                      
                      If statutcdaprgasita = 1 And statutstearsa = 0 And rest_de_sosit > 0 Then
                      
                      
                      ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 18) = ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 6) - ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 10)
                      
                      
                      restsosire = ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 18)
                      
                      MsgBox "restsosire=" & restsosire
                      
                      maxintrare = Application.WorksheetFunction.Min(restsosire, valstoc)
                      
                      
                      
                                         MsgBox "nr cda1=" & nrcdaprgasita
                      
                                        If restsosire > 0 Then
                                        
                                        
                                        'msgbox "res >0 nr cda1=" & nrcdaprgasita
                                        
                                        restsosire = restsosire - maxintrare
                                        
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 18) = restsosire
                                        
                                        ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 10) = maxintrare + ThisWorkbook.Sheets("comenzi detaliate").Cells(randulprgasita, 10)
                                        
                                        valstoc = valstoc - maxintrare
                                        
                                        ThisWorkbook.Sheets("stoc piese").Cells(liniestoc, 4) = valstoc
                                        
                                        
                                         'actualizare iesire piese
                                         
                                         linieiesirepiese = ThisWorkbook.Sheets("iesire piese").Cells(1, 1) + 2
                                        
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 2) = ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese - 1, 2) + 1
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 3) = intrarecodpiesa.Text
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 4) = maxintrare
                                         ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 5) = denumireintrare.Text
                                         
                                         'nume client
                                         
                                         poznumeclient = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzisimplu").Range("B:B"), nrcdaprgasita, ThisWorkbook.Sheets("comenzisimplu").Range("b:b"))
                                         
                                         
                                         
                                         'msgbox "poz nume client=" & poznumeclient
                                         'msgbox "nr comanda=" & nrcdaprgasita
                                         
                                         
                                           Dim rangeclient As Range
                                         
                                           Set rangeclient = ThisWorkbook.Sheets("comenzisimplu").Range("B1")
                                         
                                            Set rangeclient = Columns(2).Find(What:=nrcdaprgasita, After:=rangeclient, _
                                            LookIn:=xlValues, LookAt:=xlWhole, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                            MatchCase:=True)
                                         
                                         
                                         'msgbox "range client row=" & rangeclient.Row
                                         
                                         Dim pretvanzintrare1 As Double
                                         pretvanzintrare1 = pretvanzareintrare.Value
                                         
                                         
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 6) = ThisWorkbook.Sheets("comenzisimplu").Cells(rangeclient.Row, 3)
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 7) = ThisWorkbook.Sheets("comenzisimplu").Cells(rangeclient.Row, 4)
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 8) = pretvanzintrare1
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 10) = nrcdaprgasita
                                            ThisWorkbook.Sheets("iesire piese").Cells(linieiesirepiese, 11) = CDate(Now())
                                            
                                            
                                            
                                        
                                        End If
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      End If
                      
      
              

             'actualizare stoc, rest de sosit, iesiri daca dt este validat


        If valstoc = 0 Then
        i = iLoop + 1
        End If
        
        
        
        'actualizare statut comanda de livrat
        
        nr_pr_cerute_cda = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), nrcdaprgasita, ThisWorkbook.Sheets("comenzi detaliate").Range("f:f"))
        nr_pr_sosite_cda = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), nrcdaprgasita, ThisWorkbook.Sheets("comenzi detaliate").Range("j:j"))
                
                
                
                If nr_pr_cerute_cda = nr_pr_sosite_cda And ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 15) <> 1 And ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 25) <> 1 Then
                
                MsgBox "comanda " & nrcdaprgasita & " este de livrat!"
                
                
            
                'actualizare statut de livrat si numar de piese sosite
                ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 15) = 1
                ThisWorkbook.Sheets("comenzisimplu").Cells(nrcdaprgasita + 1, 24) = Application.WorksheetFunction.SumIf(ThisWorkbook.Sheets("comenzi detaliate").Range("c:c"), nrcdaprgasita, ThisWorkbook.Sheets("comenzi detaliate").Range("j:j"))
                 
                
                
                
                
                End If
                
                
                Next i





  End If



'sfr cautare comenzi validate care solicita pr intrata











' toate informatiile sunt resetate

                     intrarecodpiesa.Text = ""
                     denumireintrare.Text = ""
                     cantitateintrare.Text = ""
                     furnizorintrarepiese.Text = ""
                     pretachizitieintrare.Text = ""
                     pretvanzareintrare.Text = ""
'sfr  toate informatiile sunt resetate



End Sub

Private Sub cantitateintrare_Change()


If Not IsNumeric(cantitateintrare.Text) Or cantitateintrare.Text = "0" Then
cantitateintrare.Text = ""

End If



End Sub

Private Sub denumireintrare_Change()

End Sub

Private Sub furnizorintrarepiese_Change()

End Sub

Private Sub intrarecodpiesa_Change()

End Sub

Private Sub pretachizitieintrare_Change()


If pretachizitieintrare.Value <= 0 Or Not IsNumeric(pretachizitieintrare.Text) Then
pretachizitieintrare.Text = ""

End If


If IsNumeric(pretachizitieintrare.Value) And IsNumeric(adaospiese.Value) Then

pretvanzareintrare.Value = Application.WorksheetFunction.Round(pretachizitieintrare.Value * adaospiese.Value / 100 + pretachizitieintrare.Value, 2)



End If





End Sub

Private Sub pretvanzareintrare_Change()


If pretvanzareintrare.Value <= 0 Or Not IsNumeric(pretvanzareintrare.Text) Then
pretvanzareintrare.Text = ""

End If


If IsNumeric(pretachizitieintrare.Value) And IsNumeric(adaospiese.Value) Then

pretvanzareintrare.Value = Application.WorksheetFunction.Round(pretachizitieintrare.Value * adaospiese.Value / 100 + pretachizitieintrare.Value, 2)



End If




End Sub

Private Sub UserForm_Activate()





adaospiese.Text = ThisWorkbook.Sheets("Intrare piese").Cells(1, 12)





End Sub

Private Sub UserForm_Click()








End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)


intraridepiese.Hide

Aplicatie.Show vbModeless



End Sub

    
