Attribute VB_Name = "Module1"
Function GrossUp(SalNetTarget As Double, Scutire As Double, Lim7prc As Double, _
        RataMed As Double, RataFS As Double, SalMediu As Double, FacilProg As Boolean)
    
    Application.Volatile
    Dim Test As Double
        
    Best = Application.WorksheetFunction.Round(SalNetTarget / (1 - RataMed - RataFS), 2)
    
    If FacilProg Then
        Worst = Application.WorksheetFunction.Round((SalNetTarget + Lim7prc * 0.07 + (SalMediu * 2 - Lim7prc) * 0.18) / (1 - RataMed - RataFS), 2)
    Else
        Worst = Application.WorksheetFunction.Round(SalNetTarget / (1 - RataMed - RataFS - 0.18), 2)
    End If
    
    Result = 0
    Do While Result <> SalNetTarget
        Test = Application.WorksheetFunction.Round((Best + Worst) / 2, 2)
        Result = SalNet(Test, Scutire, Lim7prc, RataMed, RataFS, SalMediu, FacilProg)
        If Result < SalNetTarget Then
            Best = Test
        Else
            Worst = Test
        End If
    Loop
    
   GrossUp = Test

End Function

Private Function SalNet(SalBrutT As Double, _
                ScutireT As Double, _
                Lim7prcT As Double, _
                RataMedT As Double, _
                RataFST As Double, _
                SalMediuT As Double, _
                FacilProgT As Boolean)
    
    With Application.WorksheetFunction
        Med = .Round(SalBrutT * RataMedT, 2)
        FS = .Round(SalBrutT * RataFST, 2)
        
        If FacilProgT Then
            Impozabil = .Round(.Min(SalMediuT * 2, SalBrutT), 2)
        Else
            Impozabil = .Round(SalBrutT - Med - FS - ScutireT, 2)
        End If
        
        Impozit = .Round(.Min(Lim7prcT, Impozabil) * 0.07 + .Max(0, Impozabil - Lim7prcT) * 0.18, 2)
        SalNet = .Round(SalBrutT - Med - FS - Impozit, 2)
    End With
    
End Function

