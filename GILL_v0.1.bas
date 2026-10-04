Attribute VB_Name = "GILL_V0_1"
'==========================================================================================
' GILL - General Insurance Lambda Library
' Version 0.1
' for use in ACTIVE workbook, so it doesn't need to be resaved as an .xlsm
' Requires Excel 2024+
'==========================================================================================

'-----------------------------GLOBAL DEFINITIONS---------------------------------------
Option Explicit

Public Const GILL_VERSION As String = "0.1"
Public Const GILL_SHEET As String = "GILL_V" & GILL_VERSION
Public Const GILL_HELP_URL As String = "https://github.com/BayesOnATrueStory/GILL#function-reference"
Private Const GILL_TAG As String = "GILL"   'written into each Name's Comment
Private Const GILL_SEP As String = "|"

'-----------------------------PUBLIC MACROS------------------------------------------------
' 1. GILL_Install       (install the 14 built-in functions)
' 2. GILL_InstallFromFile (install from a pipe-delimited definitions file)
' 3. GILL_SelfTest      (evaluate all 14 functions' calculations against known answers)
' 4. GILL_Export        (write current functions to a text file - helpful for customization)
' 5. GILL_Remove        (delete only functions that this module created)
' 6. GILL_Help          (open online reference for these functions)
'
' Note: both Install and InstallFromFile are routed via the private InstallCore.
' All public subs below have no arguments, so that they can display when hitting Alt+F8
' (i.e. the macro list under the Developer tab)

'This/InstallFromFile are wrappers for the install engine below
Public Sub GILL_Install()
    GILL_InstallCore ""
End Sub

Public Sub GILL_InstallFromFile()
    Dim p As Variant
    p = Application.GetOpenFilename("GILL definitions (*.txt),*.txt", , "Select a GILL definitions file")
    If VarType(p) = vbBoolean Then Exit Sub
    GILL_InstallCore CStr(p)
End Sub

'How robust is this set of functions?
Public Sub GILL_SelfTest() 'see more on self test helper below
    Dim wb As Workbook, ws As Worksheet
    Dim tests As Collection, t As Variant, got As Variant
    Dim i As Long, pass As Long, fail As Long, tol As Double
    Dim detail As String
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    
    Set tests = GILL_Tests()   ' each record: function, formula, expected, what it checks
    
    On Error GoTo cleanup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ws = wb.Worksheets.Add
    GILL_SeedTriangle ws
    
    For i = 1 To tests.Count
        GILL_Put ws, "H" & i, CStr(tests(i)(1))
    Next i
    
    For i = 1 To tests.Count
        t = tests(i)
        tol = IIf(Abs(CDbl(t(2))) > 1, Abs(CDbl(t(2))) * 0.0000000001, 0.0000000001)
        
        If Not ws.Range("H" & i).HasFormula Then
            fail = fail + 1
            detail = detail & "FAIL  " & t(0) & " - " & t(3) & " -> Excel rejected formula entry" & vbLf
        ElseIf IsError(ws.Range("H" & i).Value) Then
            fail = fail + 1
            detail = detail & "FAIL  " & t(0) & " - " & t(3) & " -> " & ws.Range("H" & i).Text & vbLf
        Else
            got = ws.Range("H" & i).Value
            If IsArray(got) Then got = got(1, 1) ' Handle spilled dynamic arrays
            
            If IsNumeric(got) And IsNumeric(t(2)) Then
                If Abs(CDbl(got) - CDbl(t(2))) <= tol Then
                    pass = pass + 1
                Else
                    fail = fail + 1
                    detail = detail & "FAIL  " & t(0) & " - " & t(3) & _
                             "  expected " & t(2) & " got " & got & vbLf
                End If
            Else
                fail = fail + 1
                detail = detail & "FAIL  " & t(0) & " - " & t(3) & " -> Non-numeric output: " & got & vbLf
            End If
        End If
    Next i

cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    Debug.Print "-- GILL v" & GILL_VERSION & ": " & pass & " passed, " & fail & " failed --"
    Debug.Print detail
    
    If fail = 0 Then
        MsgBox "All " & pass & " checks passed." & vbLf & vbLf & GILL_Count() & _
               " functions verified against independently computed answers.", _
               vbInformation, "GILL self test - PASS"
    Else
        MsgBox pass & " passed, " & fail & " failed." & vbLf & vbLf & detail & vbLf & _
               "Full log: close this box, press Alt+F11 to open the VBA editor, then Ctrl+G.", _
               vbExclamation, "GILL self test - FAIL"
    End If
End Sub
'Allows user to remove all GILL UDFs
Public Sub GILL_Remove()
    Dim wb As Workbook, n As Name, kicked As Long, i As Long
    Dim cmt As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If MsgBox("Remove the GILL functions from """ & wb.Name & """?" & vbLf & vbLf & _
              "Any formula still calling them will return #NAME?." & vbLf & _
              "Your own defined names are not affected.", _
              vbQuestion + vbYesNo, "GILL remove") <> vbYes Then Exit Sub
    For i = wb.names.Count To 1 Step -1
        Set n = wb.names(i)
        cmt = ""
        On Error Resume Next
        cmt = n.Comment
        On Error GoTo 0
        If InStr(1, cmt, GILL_TAG, vbTextCompare) = 1 Then
            On Error Resume Next
            n.Delete
            If Err.Number = 0 Then kicked = kicked + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i
    On Error Resume Next
    Application.DisplayAlerts = False
    If wb.Worksheets.Count > 1 Then wb.Worksheets(GILL_SHEET).Delete
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0
    MsgBox kicked & " function(s) removed.", vbInformation, "GILL removed"
End Sub

'Current definitions sent to a .txt in a way that's compatible with GILL_InstallFromFile
Public Sub GILL_Export()
    Dim wb As Workbook, n As Name, p As String, ff As Integer, cnt As Long
    Dim body As String, cmt As String, pos As Long
    Dim pick As Variant

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    ' Default: next to the workbook, or Documents if unsaved / on SharePoint / OneDrive
    p = wb.Path
    If Len(p) = 0 Or Left$(LCase$(p), 4) = "http" Then
        p = Environ$("USERPROFILE") & "\Documents"
    End If
    p = p & "\GILL_export.txt"

    If MsgBox("Save to a custom location?" & vbLf & vbLf & _
              "No saves to:" & vbLf & p, _
              vbYesNo + vbQuestion, "GILL export") = vbYes Then
        pick = Application.GetSaveAsFilename( _
                   InitialFileName:=p, _
                   FileFilter:="Text files (*.txt),*.txt", _
                   Title:="Save GILL export as")
        If VarType(pick) = vbBoolean Then Exit Sub   ' cancelled
        p = CStr(pick)
    End If

    On Error GoTo ExportError
    ff = FreeFile
    Open p For Output As #ff

    Print #ff, "# GILL - General Insurance Lambda Library v" & GILL_VERSION
    Print #ff, "# Exported " & Format$(Now, "yyyy-mm-dd hh:nn") & " from " & wb.Name
    Print #ff, "# Format: NAME" & GILL_SEP & "comment" & GILL_SEP & "=LAMBDA(...)"
    Print #ff, "# Re-import with GILL_InstallFromFile. Lines starting with # are ignored."

    For Each n In wb.names
        cmt = ""
        On Error Resume Next
        cmt = n.Comment
        On Error GoTo ExportError

        If InStr(1, cmt, GILL_TAG, vbTextCompare) = 1 Then
            pos = InStr(cmt, "|")
            If pos > 0 Then
                body = Trim$(Mid$(cmt, pos + 1))
            Else
                body = cmt
            End If
            Print #ff, n.Name & GILL_SEP & body & GILL_SEP & n.RefersTo
            cnt = cnt + 1
        End If
    Next n

    Close #ff
    MsgBox cnt & " definition(s) written to:" & vbLf & p & vbLf & vbLf & _
           "You can edit it in any text editor, then re-install with GILL_InstallFromFile.", _
           vbInformation, "GILL export"
    Exit Sub

ExportError:
    If ff > 0 Then Close #ff
    MsgBox "Could not export definitions: " & Err.Description, vbCritical, "GILL export error"
End Sub

Public Sub GILL_Help() 'Redirects to a URL; too big for a MsgBox
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink Address:=GILL_HELP_URL, NewWindow:=True
    If Err.Number <> 0 Then MsgBox GILL_HELP_URL, vbInformation, "GILL reference"
End Sub



'-----------------------------FUNCTIONS------------------------------------------------

'the meat of the .bas! default installation is this unless a file is specified.
' For 3 categories (pricing, reserving, reinsurance) plus 1 freebie (XINTERP)

'-----------------------------FUNCTION REGISTRY---------------------------------------------

' For each function, assign name + description + formula + probe. Then GILL_Count will swoop
' in to tally this.
Public Function GILL_Registry() As Variant
    ' The probe is a call that must return a number once the function is installed;
    ' GILL_Install uses it to check each one works. f gets its own statement because
    ' VBA allows at most 25 line continuations per statement.
    Dim gillcol As New Collection
    Dim gillreg() As String
    Dim f As String, i As Long, j As Long

    'XINTERP
    f = "LAMBDA(x,x_known,y_known,[method]," & _
        "LET(m,IF(ISOMITTED(method),""linear"",LOWER(method))," & _
        "xr,TOCOL(x_known)," & _
        "yr,TOCOL(y_known)," & _
        "IF(ROWS(xr)<>ROWS(yr),NA()," & _
        "LET(ok,TOCOL(ISNUMBER(x_known))*TOCOL(ISNUMBER(y_known))," & _
        "xf,FILTER(xr,ok=1)," & _
        "yf,FILTER(yr,ok=1)," & _
        "xk,SORT(xf)," & _
        "yk,SORTBY(yf,xf)," & _
        "n,ROWS(xk)," & _
        "MAP(x,LAMBDA(xi," & _
        "IF(NOT(ISNUMBER(xi)),NA()," & _
        "LET(p,MATCH(xi,xk,1)," & _
        "IF(ISNA(p),NA()," & _
        "IF(p=n,IF(xi=INDEX(xk,n),INDEX(yk,n),NA())," & _
        "LET(x_lo,INDEX(xk,p),x_hi,INDEX(xk,p+1),y_lo,INDEX(yk,p),y_hi,INDEX(yk,p+1)," & _
        "IF(x_hi=x_lo,y_lo," & _
        "LET(f,(xi-x_lo)/(x_hi-x_lo)," & _
        "SWITCH(m,""linear"",y_lo+(y_hi-y_lo)*f,""log"",IF(OR(y_lo<=0,y_hi<=0),NA(),y_lo*(y_hi/y_lo)^f),""nearest"",IF(f<0.5,y_lo,y_hi),NA()))))))))))))))"
    gillcol.Add Array("XINTERP", "Interpolates a value from known data points using linear, log-linear, or nearest-neighbor methods; auto-sorts the lookup table; skips blank or text rows", f, "XINTERP(2.5,{1;2;3;4},{10;20;30;40})")

    'RATE_INDEX - note this was just SCAN at first, but I wanted to make it less forgiving
    f = "LAMBDA(changes,[base]," & _
        "LET(" & _
            "b,IF(ISOMITTED(base),1,base)," & _
            "bad,SUM(MAP(changes,LAMBDA(c,IF(ISNUMBER(c),0,1))))," & _
            "IF(OR(NOT(ISNUMBER(b)),bad>0)," & _
                "MAP(changes,LAMBDA(c,NA()))," & _
                "SCAN(b,changes,LAMBDA(cum,chg,cum*(1+chg)))" & _
            ")" & _
        ")" & _
    ")"
    gillcol.Add Array("RATE_INDEX", "Compiles cumulative rate level index from a history of rate changes", f, "INDEX(RATE_INDEX({0.05;0.08}),2)")

    'CRED_Z
    f = "LAMBDA(n,[method],[k]," & _
        "LET(" & _
            "m,IF(ISOMITTED(method),""lf"",LOWER(method))," & _
            "kv,IF(ISOMITTED(k),1082,k)," & _
            "IF(OR(NOT(ISNUMBER(kv)),kv<=0),NA()," & _
                "MAP(n,LAMBDA(nx," & _
                    "IF(OR(NOT(ISNUMBER(nx)),nx<0),NA()," & _
                        "SWITCH(m,""lf"",MIN(1,SQRT(nx/kv)),""buhlmann"",nx/(nx+kv),NA())" & _
                    ")" & _
                "))" & _
            ")" & _
        ")" & _
    ")"
    gillcol.Add Array("CRED_Z", "Calculates Limited Fluctuation or Buhlmann credibility Z factor", f, "CRED_Z(300)")

    'LAYER_LOSS
    f = "LAMBDA(el,attach,layer_lim,ilf_lims,ilf_vals,[method]," & _
        "LET(" & _
            "m,IF(ISOMITTED(method),""linear"",LOWER(method))," & _
            "exhaust,attach+layer_lim," & _
            "base,XINTERP(MIN(ilf_lims),ilf_lims,ilf_vals)," & _
            "fa,XINTERP(attach,ilf_lims,ilf_vals,m)," & _
            "fe,XINTERP(exhaust,ilf_lims,ilf_vals,m)," & _
            "IF(OR(ISNA(fa),ISNA(fe),NOT(ISNUMBER(base)),base<=0),NA()," & _
                "MAP(el,LAMBDA(e,IF(OR(NOT(ISNUMBER(e)),e<0),NA(),e*(fe-fa)/base)))" & _
            ")" & _
        ")" & _
    ")"
    gillcol.Add Array("LAYER_LOSS", "Calculates expected layer loss using ILFs; blank rows in the ILF table are ignored", f, "LAYER_LOSS(500000,1000000,5000000,{1000000;2000000;5000000;10000000},{1;1.25;1.58;1.75})")

    'MBBEFD
    f = "LAMBDA(d,b,g," & _
        "IF(OR(NOT(ISNUMBER(b)),NOT(ISNUMBER(g)),b<0,g<1),NA()," & _
            "MAP(d,LAMBDA(dx," & _
            "IF(NOT(ISNUMBER(dx)),NA()," & _
            "LET(x,MAX(0,MIN(1,dx))," & _
            "IF(x=0,0,IF(x=1,1," & _
            "IF(OR(b=0,g=1),x," & _
            "IF(b=1,LN(1+(g-1)*x)/LN(g)," & _
            "IF(ABS(b*g-1)<0.000000001,(1-b^x)/(1-b)," & _
            "LN(((g-1)*b+(1-g*b)*b^x)/(1-b))/LN(g*b))))))))))))"
    gillcol.Add Array("MBBEFD", "Returns the Swiss Re / Bernegger exposure curve value G(d)", f, "MBBEFD(0.5,2,3)")

    'ON_LEVEL
    f = "LAMBDA(pol_eff,rate_dates,rate_chgs,[as_of]," & _
        "LET(" & _
            "ad,IF(ISOMITTED(as_of),TODAY(),as_of)," & _
            "dr,TOCOL(rate_dates)," & _
            "cr,TOCOL(rate_chgs)," & _
            "bad,SUM(MAP(dr,LAMBDA(v,IF(ISNUMBER(v),0,1))))+SUM(MAP(cr,LAMBDA(v,IF(ISNUMBER(v),0,1))))," & _
            "IF(OR(ROWS(dr)<>ROWS(cr),bad>0)," & _
                "MAP(pol_eff,LAMBDA(v,NA()))," & _
                "LET(" & _
                    "rd,SORT(dr)," & _
                    "rc,SORTBY(cr,dr)," & _
                    "idx,RATE_INDEX(rc,1)," & _
                    "lvl,LAMBDA(dt,LET(p,SUMPRODUCT((rd<=dt)*1),IF(p=0,1,INDEX(idx,p))))," & _
                    "cur,lvl(ad)," & _
                    "MAP(pol_eff,LAMBDA(eff,IF(NOT(ISNUMBER(eff)),NA(),cur/lvl(eff))))" & _
                ")" & _
            ")" & _
        ")" & _
    ")"
    gillcol.Add Array("ON_LEVEL", "Applies on-level factors using historical rate change arrays; auto-sorts the rate history", f, "ON_LEVEL(DATE(2021,7,1),DATE({2023;2022},1,1),{0.08;0.05},DATE(2025,1,1))")

    'ALAE_LOAD
    f = "LAMBDA(indem,method,param,[cap]," & _
        "LET(" & _
            "m,LOWER(method)," & _
            "c,IF(ISOMITTED(cap),1E+99,cap)," & _
            "MAP(indem,LAMBDA(i," & _
                "IF(NOT(ISNUMBER(i)),NA()," & _
                    "LET(ind,MAX(0,i)," & _
                    "SWITCH(m,""pro_rata"",ind*param,""flat"",param,""capped"",MIN(ind*param,c),NA()))" & _
                ")" & _
            "))" & _
        ")" & _
    ")"
    gillcol.Add Array("ALAE_LOAD", "Applies ALAE loading to indemnity losses (pro_rata, flat, capped)", f, "ALAE_LOAD(0,""flat"",5000)")

    'TRIANGLE
    f = "LAMBDA(origin,dev,values,[cum]," & _
        "LET(" & _
            "is_cum,IF(ISOMITTED(cum),FALSE,cum)," & _
            "uo,SORT(UNIQUE(FILTER(origin,ISNUMBER(origin))))," & _
            "ud,SORT(UNIQUE(FILTER(dev,ISNUMBER(dev))))," & _
            "md,MAP(uo,LAMBDA(k,MAXIFS(dev,origin,k)))," & _
            "body,MAKEARRAY(ROWS(uo),ROWS(ud),LAMBDA(r,c," & _
                "LET(o,INDEX(uo,r),d,INDEX(ud,c)," & _
                "IF(d>INDEX(md,r),""""," & _
                "IF(is_cum,SUMIFS(values,origin,o,dev,""<=""&d),SUMIFS(values,origin,o,dev,d))))))," & _
            "VSTACK(HSTACK(""Origin"",TOROW(ud)),HSTACK(uo,body))" & _
        ")" & _
    ")"
    gillcol.Add Array("TRIANGLE", "Converts tabular loss data (origin, dev, value) into a 2D triangle array. Requires ranges as inputs.", f, "INDEX(TRIANGLE($A$2:$A$7,$B$2:$B$7,$C$2:$C$7),2,2)")

    'LDF_SELECT
    f = "LAMBDA(factors,[method],[n_periods],[weights]," & _
        "LET(" & _
            "m,IF(ISOMITTED(method),""simple"",LOWER(method))," & _
            "fa,TOCOL(factors)," & _
            "keep,TOCOL(ISNUMBER(factors))," & _
            "fc,FILTER(fa,keep)," & _
            "ct,ROWS(fc)," & _
            "IF(ct=0,NA()," & _
                "LET(n,IF(ISOMITTED(n_periods),ct,MIN(MAX(1,n_periods),ct))," & _
                "ft,TAKE(fc,-n)," & _
                "SWITCH(m,""simple"",AVERAGE(ft)," & _
                """medial"",IF(n<=2,AVERAGE(ft),AVERAGE(DROP(DROP(SORT(ft),1),-1)))," & _
                """volume"",IF(ISOMITTED(weights),NA()," & _
                    "LET(wa,TOCOL(weights)," & _
                    "IF(ROWS(wa)<>ROWS(fa),NA()," & _
                        "LET(wt,TAKE(FILTER(wa,keep),-n)," & _
                        "wn,TAKE(FILTER(TOCOL(ISNUMBER(weights)),keep),-n)," & _
                        "IF(SUM(--wn)<ROWS(wn),NA()," & _
                        "IF(SUM(wt)=0,NA(),SUM(ft*wt)/SUM(wt)))))))," & _
                "NA())))))"
    gillcol.Add Array("LDF_SELECT", "Selects LDFs using simple, medial, or volume-weighted averages; skips blanks and text, so a column from a factor triangle works as-is", f, "LDF_SELECT({1.892;1.756;2.104;1.834;1.901},""medial"")")

'I've been eating VBA for lunch and dinner. I'm following a macro diet.

    'ULTIMATE
    f = "LAMBDA(reported,cdf,[method],[elr],[premium]," & _
        "LET(" & _
            "m,IF(ISOMITTED(method),""cl"",LOWER(method))," & _
            "bad,SUM(MAP(reported,LAMBDA(r,IF(ISNUMBER(r),0,1))))," & _
            "IF(bad>0,MAP(reported,LAMBDA(r,NA()))," & _
                "LET(pu,1-1/cdf," & _
                "SWITCH(m,""cl"",reported*cdf,""bf""," & _
                "IF(OR(ISOMITTED(elr),ISOMITTED(premium)),NA(),reported+elr*premium*pu)," & _
                """cc"",IF(ISOMITTED(premium),NA()," & _
                    "LET(used,premium/cdf,dn,SUM(used)," & _
                    "IF(dn=0,NA(),reported+SUM(reported)/dn*premium*pu))),NA())))))"
    gillcol.Add Array("ULTIMATE", "Consolidates chain-ladder, Bornhuetter-Ferguson, and Cape Cod reserving methods", f, "ULTIMATE(8500000,1.25,""bf"",0.65,15000000)")

    'AGG_ERODE
    f = "LAMBDA(losses,retention,limit,aggregate,[agg_type]," & _
        "LET(" & _
            "typ,IF(ISOMITTED(agg_type),""ded"",LOWER(agg_type))," & _
            "bad,SUM(MAP(losses,LAMBDA(l,IF(ISNUMBER(l),0,1))))," & _
            "IF(OR(NOT(OR(typ=""ded"",typ=""lim"")),bad>0)," & _
                "MAP(losses,LAMBDA(l,NA()))," & _
                "LET(" & _
                    "occ,MAP(losses,LAMBDA(l,MAX(0,MIN(l-retention,limit))))," & _
                    "cum,SCAN(0,occ,LAMBDA(a,b,a+b))," & _
                    "MAP(occ,cum,LAMBDA(o,cm," & _
                        "LET(prior,cm-o,rem,MAX(0,aggregate-prior)," & _
                        "IF(typ=""ded"",MAX(0,o-rem),MIN(o,rem))" & _
                    ")))" & _
                ")" & _
            ")" & _
        ")" & _
    ")"
    gillcol.Add Array("AGG_ERODE", "Tracks aggregate limit/deductible erosion across a sequence of losses", f, "INDEX(AGG_ERODE({4000000;6000000;8000000;5000000;10000000},3000000,2000000,5000000),4)")

    'CORRIDOR
    f = "LAMBDA(lr,attach,width,base_cess,[corr_cess]," & _
        "LET(" & _
            "cc,IF(ISOMITTED(corr_cess),0,corr_cess)," & _
            "exhaust,attach+width," & _
            "MAP(lr,LAMBDA(v," & _
                "IF(NOT(ISNUMBER(v)),NA()," & _
                    "LET(r,MAX(0,v)," & _
                    "below,MIN(r,attach)," & _
                    "within,MAX(0,MIN(r-attach,width))," & _
                    "above,MAX(0,r-exhaust)," & _
                    "IF(r=0,base_cess,(below*base_cess+within*cc+above*base_cess)/r))" & _
                ")" & _
            "))" & _
        ")" & _
    ")"
    gillcol.Add Array("CORRIDOR", "Blends standard and inside-corridor cession percentages", f, "CORRIDOR(0.75,0.7,0.1,0.5,0)")

    'REINST_PREM
    f = "LAMBDA(losses,ret,lim,base_prem,n_reinst,[rate],[type]," & _
        "LET(" & _
            "rt,IF(ISOMITTED(rate),1,rate)," & _
            "typ,IF(ISOMITTED(type),""pro_rata"",LOWER(type))," & _
            "bad,SUM(MAP(losses,LAMBDA(l,IF(ISNUMBER(l),0,1))))," & _
            "IF(OR(NOT(ISNUMBER(lim)),lim<=0,bad>0),NA()," & _
                "LET(" & _
                    "maxc,lim*(1+n_reinst)," & _
                    "occ,MAP(losses,LAMBDA(l,MAX(0,MIN(l-ret,lim))))," & _
                    "tot,REDUCE(0,occ,LAMBDA(cm,o,LET(rem,MAX(0,maxc-cm),cm+MIN(o,rem))))," & _
                    "xs,MAX(0,tot-lim)," & _
                    "SWITCH(typ,""pro_rata"",xs/lim*rt*base_prem,""flat"",MIN(n_reinst,CEILING(xs/lim,1))*rt*base_prem,NA())" & _
                ")" & _
            ")" & _
        ")" & _
    ")"
    gillcol.Add Array("REINST_PREM", "Calculates sequential reinstatement premium (pro rata or flat)", f, "REINST_PREM({8000000;12000000;7000000;15000000},5000000,5000000,500000,2)")

    'XOL_RECOV
    f = "LAMBDA(losses,retention,limit,[loss_cap]," & _
        "MAP(losses,LAMBDA(l," & _
            "IF(NOT(ISNUMBER(l)),NA()," & _
                "LET(cp,IF(ISOMITTED(loss_cap),l,MIN(l,loss_cap))," & _
                "MAX(0,MIN(cp-retention,limit)))" & _
            ")" & _
        "))" & _
    ")"
    gillcol.Add Array("XOL_RECOV", "Calculates per-occurrence excess of loss recovery", f, "XOL_RECOV(8500000,5000000,5000000)")

    ReDim gillreg(1 To gillcol.Count, 1 To 4)
    For i = 1 To gillcol.Count
        For j = 0 To 3
            gillreg(i, j + 1) = gillcol(i)(j)
        Next j
    Next i
    GILL_Registry = gillreg
End Function

Public Function GILL_Count() As Long
    GILL_Count = UBound(GILL_Registry(), 1)
End Function

'-----------------------------INSTALLATION ENGINE ---------------------------------------

Private Sub GILL_InstallCore(ByVal filePath As String)
    ' filePath empty -> install the built-in set
    ' filePath given -> validate, then install from it.
    '                   if any structural issue's spotted, report line# then stop.
    Dim wb As Workbook, installed As Long, total As Long, problem As String, src As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Open a workbook first.", vbExclamation, "GILL"
        Exit Sub
    End If
    If Len(filePath) > 0 Then
        If Not GILL_FileIsValid(filePath, problem) Then
            MsgBox "That file cannot be imported." & vbLf & vbLf & problem, _
                   vbCritical, "GILL - invalid definitions file"
            Exit Sub
        End If
        installed = GILL_InstallFile(wb, filePath, problem, total)
        src = "file: " & filePath
    Else
        installed = GILL_InstallBuiltIn(wb, problem)
        total = GILL_Count()
        src = "built-in set (v" & GILL_VERSION & ")"
    End If
    GILL_WriteInfoSheet wb, installed, src
    If Len(problem) = 0 Then
        MsgBox "GILL v" & GILL_VERSION & " installed successfully." & vbLf & vbLf & _
               installed & " of " & total & " functions are available in """ & wb.Name & """." & vbLf & _
               "Try:  =XOL_RECOV(8500000, 5000000, 5000000)", _
               vbInformation, "GILL installed"
    Else
        MsgBox installed & " of " & total & " functions installed and working." & vbLf & vbLf & problem, _
               vbExclamation, "GILL - partial install"
    End If
End Sub

Private Function GILL_InstallBuiltIn(wb As Workbook, ByRef problem As String) As Long
    Dim reg As Variant, bad As Long, names As String
    reg = GILL_Registry()
    bad = GILL_AddAll(wb, reg, True, names)
    If bad > 0 Then bad = GILL_AddAll(wb, reg, False, names)
    If bad > 0 Then
        problem = "Not working: " & names & vbLf & vbLf & _
                  "Usual causes: a definition was edited and no longer parses, or this " & _
                  "version of Excel lacks a function it uses. A function that calls " & _
                  "another (LAYER_LOSS calls XINTERP, ON_LEVEL calls RATE_INDEX) also " & _
                  "fails when that one does." & vbLf & vbLf & _
                  "Run GILL_SelfTest for test-by-test detail."
    End If
    GILL_InstallBuiltIn = GILL_Count() - bad
End Function

'---Text Validation to Enable File Installation---
' Format, one function per line:  NAME|comment|=LAMBDA(...)
' Blank lines and lines starting with # are ignored.
Private Function GILL_FileIsValid(ByVal filePath As String, ByRef problem As String) As Boolean
    
    Dim fileNum As Integer, ln As String, parts() As String
    Dim lineNo As Long, found As Long
    problem = ""
    
    On Error GoTo Failed
    
    'First things first - does the file exist?
    If Len(Dir(filePath)) = 0 Then
        problem = "File not found."
        Exit Function
    End If
    
    fileNum = FreeFile 'this points at the line number of the file present
    On Error GoTo Failed
    Open filePath For Input As #fileNum
    
    'Error checker
Do While Not EOF(fileNum)
        Line Input #fileNum, ln
        lineNo = lineNo + 1
        ln = Trim$(ln)
        
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            parts = Split(ln, GILL_SEP)
            If UBound(parts) <> 2 Then
                problem = "Line " & lineNo & ": expected 3 fields separated by """ & _
                          GILL_SEP & """, found " & (UBound(parts) + 1) & "."
                Close #fileNum
                Exit Function
            ElseIf Len(Trim$(parts(0))) = 0 Then
                problem = "Line " & lineNo & ": function name is blank."
                Close #fileNum
                Exit Function
            ElseIf InStr(Trim$(parts(2)), "LAMBDA(") = 0 Then
                problem = "Line " & lineNo & " (" & Trim$(parts(0)) & _
                          "): definition does not contain LAMBDA(."
                Close #fileNum
                Exit Function
            End If
            found = found + 1
        End If
    Loop
    Close #fileNum
    
    If found = 0 Then
        problem = "No definition lines found. Expected NAME" & GILL_SEP & _
                  "comment" & GILL_SEP & "=LAMBDA(...)"
        Exit Function
    End If
    
    GILL_FileIsValid = True
    Exit Function

Failed:
    If fileNum > 0 Then Close #fileNum
    problem = "Could not read the file: " & Err.Description
End Function

' This runs only after GILL_FileIsValid has passed, so we know no successive
' errors are due to malformation. Similar structure to GILL_FileIsValid
Private Function GILL_InstallFile(wb As Workbook, ByVal filePath As String, ByRef problem As String, ByRef attempted As Long) As Long
    Dim fileNum As Integer, ln As String, parts() As String
    Dim lineNo As Long, okCount As Long, rejected As String
    
    problem = ""
    attempted = 0
    fileNum = FreeFile
    
    On Error GoTo FileError
    Open filePath For Input As #fileNum
    
    Application.ScreenUpdating = False
    Do While Not EOF(fileNum)
        Line Input #fileNum, ln
        lineNo = lineNo + 1
        ln = Trim$(ln)
        
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            parts = Split(ln, GILL_SEP)
            If UBound(parts) = 2 Then
                attempted = attempted + 1
                If GILL_AddOne(wb, Trim$(parts(0)), Trim$(parts(1)), Trim$(parts(2))) Then
                    okCount = okCount + 1
                Else
                    rejected = rejected & "Line " & lineNo & " (" & Trim$(parts(0)) & _
                                "): Excel rejected the formula." & vbLf
                End If
            End If
        End If
    Loop
    Close #fileNum
    Application.ScreenUpdating = True
    
    problem = rejected
    GILL_InstallFile = okCount
    Exit Function

FileError:
    Close #fileNum
    Application.ScreenUpdating = True
    problem = "Could not read the file: " & Err.Description
End Function

'-----------------------------BACKWARDS COMPATIBILITY--------------------------------------

' Post-2007 functions are stored with an _xlfn. prefix that Excel hides in cell UI.
' Older VBA parsers accept only prefixed functions from newer workbooks; newer ones
' accept both. So the logic below goes with "try prefixed first, fall back to bare".

Private Function GILL_Prefix(ByVal f As String) As String
    Dim xlfn_funs As Variant, xlws_funs As Variant, i As Long
    
    xlfn_funs = Array("UNIQUE", "LAMBDA", "LET", "MAP", "SCAN", "REDUCE", "MAKEARRAY", _
                      "ISOMITTED", "SWITCH", "TAKE", "DROP", "TOCOL", "TOROW", "MAXIFS")
    For i = LBound(xlfn_funs) To UBound(xlfn_funs)
        If InStr(1, f, "_xlfn." & xlfn_funs(i), vbBinaryCompare) = 0 Then
            f = Replace(f, xlfn_funs(i), "_xlfn." & xlfn_funs(i))
        End If
    Next i
    
    xlws_funs = Array("SORTBY", "SORT", "FILTER", "VSTACK", "HSTACK")
    For i = LBound(xlws_funs) To UBound(xlws_funs)
        If InStr(1, f, "_xlfn._xlws." & xlws_funs(i), vbBinaryCompare) = 0 Then
            ' Process SORTBY before SORT to avoid substring collisions
            f = Replace(f, xlws_funs(i), "_xlfn._xlws." & xlws_funs(i))
        End If
    Next i
    
    GILL_Prefix = f
End Function

'-----------------------------OTHER NAME MANAGER FUNCTIONS--------------------------
' True only if the name exists afterwards, so a rejected formula can never
' be counted as installed.
Private Function GILL_AddOne(wb As Workbook, nm As String, cmt As String, f As String) As Boolean
    
    Dim formulatext As String
    formulatext = IIf(Left$(f, 1) = "=", f, "=" & f)
    
    On Error GoTo ErrorHandler
    
    '1: if it exists, delete it to avoid collisions. Then add to Name Manager
    On Error Resume Next
        wb.names(nm).Delete
    On Error GoTo ErrorHandler
    wb.names.Add Name:=nm, RefersTo:=formulatext
    wb.names(nm).Comment = GILL_TAG & " v" & GILL_VERSION & "|" & cmt
    GILL_AddOne = True
    Exit Function
ErrorHandler:
    GILL_AddOne = False
End Function

Private Function GILL_AddAll(wb As Workbook, reg As Variant, usePrefix As Boolean, ByRef names As String) As Long
    
    Dim i As Long, f As String
    Application.ScreenUpdating = False
    
    For i = 1 To GILL_Count()
        f = reg(i, 3)
        If usePrefix Then f = GILL_Prefix(f)
        GILL_AddOne wb, CStr(reg(i, 1)), CStr(reg(i, 2)), f
    Next i
    Application.ScreenUpdating = True
    
    GILL_AddAll = GILL_CountBroken(wb, names)
End Function

'--- Scratch sheet helpers ----------------------------------------------
Private Sub GILL_SeedTriangle(ws As Worksheet)
    ws.Range("A1:C1").Value = Array("AY", "Dev", "Paid")
    ws.Range("A2:C7").Value = Application.Transpose(Array( _
        Array(2021, 2021, 2021, 2022, 2022, 2023), _
        Array(12, 24, 36, 12, 24, 12), _
        Array(500000, 800000, 950000, 600000, 900000, 550000)))
    ws.Range("E1:G1").Value = Array("AY", "Lag", "CumPaid")
    ws.Range("E2:G7").Value = Application.Transpose(Array( _
        Array(2021, 2021, 2021, 2022, 2022, 2023), _
        Array(1, 2, 3, 1, 2, 1), _
        Array(100, 150, 170, 120, 160, 130)))
    ' Ranges with deliberately EMPTY cells (J3, K2, row 5) - array constants can't hold blanks
    ws.Range("J1").Value = 1.5: ws.Range("J2").Value = 2.5: ws.Range("J4").Value = 3.5
    ws.Range("K1").Value = 10: ws.Range("K3").Value = 99: ws.Range("K4").Value = 30
    ws.Range("L1:M4").Value = Application.Transpose(Array(Array(1, 2, 3, 4), Array(10, 20, 30, 40)))
    ws.Range("N1:O4").Value = Application.Transpose(Array( _
        Array(1000000, 2000000, 5000000, 10000000), Array(1, 1.25, 1.58, 1.75)))
End Sub

Private Sub GILL_Put(ws As Worksheet, addr As String, ByVal f As String)
    On Error Resume Next
    ws.Range(addr).Formula2 = "=" & f
    If Err.Number <> 0 Then
        Err.Clear
        ws.Range(addr).Formula = "=" & f
    End If
    On Error GoTo 0
End Sub

Private Function GILL_CountBroken(wb As Workbook, ByRef names As String) As Long
    ' Writes each function's probe on a scratch sheet; counts and names the ones that fail.
    Dim ws As Worksheet, reg As Variant, i As Long, bad As Long, v As Variant, broken As Boolean
    reg = GILL_Registry()
    names = ""
    On Error GoTo cleanup
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set ws = wb.Worksheets.Add
    GILL_SeedTriangle ws
    For i = 1 To UBound(reg, 1)
        GILL_Put ws, "H" & i, "IF(ISERROR(" & reg(i, 4) & "),1,0)"
    Next i
    For i = 1 To UBound(reg, 1)
        v = ws.Range("H" & i).Value
        broken = False
        If Not ws.Range("H" & i).HasFormula Then
            broken = True
        ElseIf IsError(v) Then
            broken = True
        ElseIf v <> 0 Then
            broken = True
        End If
        If broken Then
            bad = bad + 1
            names = names & IIf(Len(names) > 0, ", ", "") & reg(i, 1)
        End If
    Next i
cleanup:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    GILL_CountBroken = bad
End Function


'--- In-workbook reference sheet ----------------------------------------
' Rebuilt on every install.
' Only documentation, so safe to delete (and goes away with GILL_Remove too)
' It's best to keep a copy for tech review, since the name manager would condense it all
' into one looooong line.

Private Sub GILL_WriteInfoSheet(wb As Workbook, ByVal installed As Long, ByVal src As String)
    Dim ws As Worksheet, r As Long
    
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(GILL_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ws.Name = GILL_SHEET

    ws.Range("B2").Value = "GILL - General Insurance Lambda Library"
    ws.Range("B2").Font.Size = 14
    ws.Range("B2").Font.Bold = True
    ws.Range("B3").Value = "Version " & GILL_VERSION & "  |  installed " & _
        Format$(Now, "yyyy-mm-dd hh:nn") & "  |  " & installed & " function(s) from " & src

    ws.Range("B5").Value = "Overview of GILL"
    ws.Range("B5").Font.Bold = True
    ws.Range("B6").Value = "The functions below are stored in this workbook as defined names."
    ws.Range("B7").Value = "You can save this file as a normal .xlsx and send it on."
    ws.Range("B8").Value = "The workbook can be saved macro-free."
    ws.Range("B9").Value = "This sheet is documentation only. Delete it whenever you like."

    ws.Range("B11").Value = "Macros (Alt+F8)"
    ws.Range("B11").Font.Bold = True
    r = 12
    ws.Cells(r, 2).Value = "GILL_Install": ws.Cells(r, 3).Value = "Install or refresh the built-in functions.": r = r + 1
    ws.Cells(r, 2).Value = "GILL_InstallFromFile": ws.Cells(r, 3).Value = "Install from your own definitions file instead.": r = r + 1
    ws.Cells(r, 2).Value = "GILL_SelfTest": ws.Cells(r, 3).Value = "Check every function against a known answer.": r = r + 1
    ws.Cells(r, 2).Value = "GILL_Export": ws.Cells(r, 3).Value = "Write the current definitions to GILL_export.txt.": r = r + 1
    ws.Cells(r, 2).Value = "GILL_Remove": ws.Cells(r, 3).Value = "Remove the functions and this sheet.": r = r + 1
    ws.Cells(r, 2).Value = "GILL_Help": ws.Cells(r, 3).Value = "Open the online reference.": r = r + 1

    ws.Cells(r + 1, 2).Value = "Functions"
    ws.Cells(r + 1, 2).Font.Bold = True
    r = r + 2
    ws.Cells(r, 2).Value = "Name": ws.Cells(r, 3).Value = "Syntax"
    ws.Cells(r, 4).Value = "Category": ws.Cells(r, 5).Value = "What it does"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "XINTERP": ws.Cells(r, 3).Value = "XINTERP(x, x_known, y_known, [method])": ws.Cells(r, 4).Value = "Pricing": ws.Cells(r, 5).Value = "Interpolates a value from known data points using linear, log-linear, or nearest-neighbor methods; auto-sorts the lookup table; skips blank or text rows": r = r + 1
    ws.Cells(r, 2).Value = "RATE_INDEX": ws.Cells(r, 3).Value = "RATE_INDEX(changes, [base])": ws.Cells(r, 4).Value = "Pricing": ws.Cells(r, 5).Value = "Cumulative rate level index from a history of rate changes.": r = r + 1
    ws.Cells(r, 2).Value = "CRED_Z": ws.Cells(r, 3).Value = "CRED_Z(n, [method], [k])": ws.Cells(r, 4).Value = "Pricing": ws.Cells(r, 5).Value = "Credibility Z - limited fluctuation or Buhlmann.": r = r + 1
    ws.Cells(r, 2).Value = "LAYER_LOSS": ws.Cells(r, 3).Value = "LAYER_LOSS(el, attach, layer_lim, ilf_lims, ilf_vals, [method])": ws.Cells(r, 4).Value = "Pricing": ws.Cells(r, 5).Value = "Expected loss to an excess layer from an ILF table.": r = r + 1
    ws.Cells(r, 2).Value = "MBBEFD": ws.Cells(r, 3).Value = "MBBEFD(d, b, g)": ws.Cells(r, 4).Value = "Pricing": ws.Cells(r, 5).Value = "Swiss Re / Bernegger exposure curve G(d).": r = r + 1
    ws.Cells(r, 2).Value = "ON_LEVEL": ws.Cells(r, 3).Value = "ON_LEVEL(pol_eff, rate_dates, rate_chgs, [as_of])": ws.Cells(r, 4).Value = "Pricing": ws.Cells(r, 5).Value = "On-level factor. Auto-sorts the rate history.": r = r + 1
    ws.Cells(r, 2).Value = "ALAE_LOAD": ws.Cells(r, 3).Value = "ALAE_LOAD(indem, method, param, [cap])": ws.Cells(r, 4).Value = "Reserving": ws.Cells(r, 5).Value = "ALAE loading - pro_rata, flat or capped.": r = r + 1
    ws.Cells(r, 2).Value = "TRIANGLE": ws.Cells(r, 3).Value = "TRIANGLE(origin, dev, values, [cum])": ws.Cells(r, 4).Value = "Reserving": ws.Cells(r, 5).Value = "Tabular (origin, dev, value) into a 2-D triangle, blank beyond each origin's latest age. Inputs must be RANGES (uses SUMIFS/MAXIFS).": r = r + 1
    ws.Cells(r, 2).Value = "LDF_SELECT": ws.Cells(r, 3).Value = "LDF_SELECT(factors, [method], [n_periods], [weights])": ws.Cells(r, 4).Value = "Reserving": ws.Cells(r, 5).Value = "LDF selection - simple, medial or volume-weighted.": r = r + 1
    ws.Cells(r, 2).Value = "ULTIMATE": ws.Cells(r, 3).Value = "ULTIMATE(reported, cdf, [method], [elr], [premium])": ws.Cells(r, 4).Value = "Reserving": ws.Cells(r, 5).Value = "Ultimate loss - chain ladder, Bornhuetter-Ferguson or Cape Cod.": r = r + 1
    ws.Cells(r, 2).Value = "AGG_ERODE": ws.Cells(r, 3).Value = "AGG_ERODE(losses, retention, limit, aggregate, [agg_type])": ws.Cells(r, 4).Value = "Reinsurance": ws.Cells(r, 5).Value = "Aggregate deductible / aggregate limit erosion.": r = r + 1
    ws.Cells(r, 2).Value = "CORRIDOR": ws.Cells(r, 3).Value = "CORRIDOR(lr, attach, width, base_cess, [corr_cess])": ws.Cells(r, 4).Value = "Reinsurance": ws.Cells(r, 5).Value = "Effective cession rate through a loss corridor.": r = r + 1
    ws.Cells(r, 2).Value = "REINST_PREM": ws.Cells(r, 3).Value = "REINST_PREM(losses, ret, lim, base_prem, n_reinst, [rate], [type])": ws.Cells(r, 4).Value = "Reinsurance": ws.Cells(r, 5).Value = "Reinstatement premium - pro rata or flat.": r = r + 1
    ws.Cells(r, 2).Value = "XOL_RECOV": ws.Cells(r, 3).Value = "XOL_RECOV(losses, retention, limit, [loss_cap])": ws.Cells(r, 4).Value = "Reinsurance": ws.Cells(r, 5).Value = "Per-occurrence excess of loss recovery.": r = r + 1

    ws.Cells(r + 1, 2).Value = "If something returns #N/A"
    ws.Cells(r + 1, 2).Font.Bold = True
    ws.Cells(r + 2, 2).Value = "That is deliberate. Text in a running-total function (AGG_ERODE,"
    ws.Cells(r + 3, 2).Value = "RATE_INDEX, ON_LEVEL, REINST_PREM) invalidates the whole result, not"
    ws.Cells(r + 4, 2).Value = "just one row. A missing required argument, an unknown method name, or"
    ws.Cells(r + 5, 2).Value = "a lookup outside a known table also return #N/A rather than a number."
    ws.Cells(r + 7, 2).Value = "Reference:"
    ws.Cells(r + 7, 3).Value = GILL_HELP_URL

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 62
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 60
    ws.Activate
    ws.Range("B2").Select
End Sub

'--- Self test helper -------------------------------------------------

Private Function GILL_Tests() As Collection
    
    ' Making a collection sorta like the GILL function collection above.
    ' Each line is function -> formula -> expected val -> what it's supposed to check
    ' Each formula returns a number.
    Dim t As New Collection
    gillt.Add Array("XINTERP", "XINTERP(2.5,{1;2;3;4},{10;20;30;40})", 25#, "linear interpolation")
    gillt.Add Array("XINTERP", "XINTERP(2.5,{3;1;4;2},{30;10;40;20})", 25#, "FIX: unsorted input now sorted")
    gillt.Add Array("CRED_Z", "CRED_Z(300)", 0.526558947624551, "limited fluctuation Z")
    gillt.Add Array("LAYER_LOSS", "LAYER_LOSS(500000,1000000,5000000,{1000000;2000000;5000000;10000000},{1;1.25;1.58;1.75})", 307000#, "ILF layer loss")
    gillt.Add Array("MBBEFD", "MBBEFD(0.5,2,3)", 0.626214255790261, "FIX: Bernegger general case")
    gillt.Add Array("MBBEFD", "MBBEFD(1,9.025,7.6906)", 1#, "FIX: G(1)=1 boundary")
    gillt.Add Array("RATE_INDEX", "INDEX(RATE_INDEX({0.05;0.08}),2)", 1.134, "cumulative rate index")
    gillt.Add Array("ON_LEVEL", "ON_LEVEL(DATE(2021,7,1),DATE({2023;2022},1,1),{0.08;0.05},DATE(2025,1,1))", 1.134, "FIX: unsorted rate history")
    gillt.Add Array("ALAE_LOAD", "ALAE_LOAD(0,""flat"",5000)", 5000#, "FIX: closed-no-pay keeps ALAE")
    gillt.Add Array("ALAE_LOAD", "ALAE_LOAD(700000,""capped"",0.15,100000)", 100000#, "capped ALAE")
    gillt.Add Array("LDF_SELECT", "LDF_SELECT({1.892;1.756;2.104;1.834;1.901},""medial"")", 1.87566666666667, "medial average")
    gillt.Add Array("LDF_SELECT", "LDF_SELECT({1.5;2.5},""medial"")", 2#, "FIX: medial on 2 points")
    gillt.Add Array("LDF_SELECT", "IF(ISNA(LDF_SELECT({1.5;2.5;3.5},""volume"")),1,0)", 1#, "FIX: volume without weights fails loudly")
    gillt.Add Array("ULTIMATE", "ULTIMATE(8500000,1.25)", 10625000#, "chain ladder")
    gillt.Add Array("ULTIMATE", "ULTIMATE(8500000,1.25,""bf"",0.65,15000000)", 10450000#, "Bornhuetter-Ferguson")
    gillt.Add Array("ULTIMATE", "IF(ISNA(ULTIMATE(8500000,1.25,""bf"")),1,0)", 1#, "FIX: BF without elr/prem fails loudly")
    gillt.Add Array("ULTIMATE", "IF(ISNA(ULTIMATE(8500000,1.25,""typo"")),1,0)", 1#, "FIX: bad method fails loudly")
    gillt.Add Array("AGG_ERODE", "INDEX(AGG_ERODE({4000000;6000000;8000000;5000000;10000000},3000000,2000000,5000000),4)", 2000000#, "aggregate deductible erosion")
    gillt.Add Array("CORRIDOR", "CORRIDOR(0.75,0.7,0.1,0.5,0)", 0.466666666666667, "loss corridor cession")
    gillt.Add Array("REINST_PREM", "REINST_PREM({8000000;12000000;7000000;15000000},5000000,5000000,500000,2)", 1000000#, "reinstatement premium")
    gillt.Add Array("XOL_RECOV", "XOL_RECOV(8500000,5000000,5000000)", 3500000#, "per-occurrence XOL recovery")
    gillt.Add Array("TRIANGLE", "INDEX(TRIANGLE($A$2:$A$7,$B$2:$B$7,$C$2:$C$7),2,2)", 500000#, "tabular data to triangle")
    gillt.Add Array("AGG_ERODE", "IF(ISNA(INDEX(AGG_ERODE({4000000;""x"";8000000;5000000;10000000},3000000,2000000,5000000),1)),1,0)", 1#, "FIX: bad row invalidates whole array, not just downstream")
    gillt.Add Array("AGG_ERODE", "INDEX(AGG_ERODE({4000000;6000000;8000000;5000000;10000000},3000000,2000000,5000000),5)", 2000000#, "clean data still correct")
    gillt.Add Array("RATE_INDEX", "IF(ISNA(INDEX(RATE_INDEX({0.05;""x"";0.03}),1)),1,0)", 1#, "FIX: SCAN poisoning in RATE_INDEX")
    gillt.Add Array("ON_LEVEL", "IF(ISNA(ON_LEVEL(DATE(2021,7,1),DATE({2022;2023},1,1),{0.05;""x""},DATE(2025,1,1))),1,0)", 1#, "FIX: poisoning inherited via RATE_INDEX")
    gillt.Add Array("ULTIMATE", "IF(ISNA(ULTIMATE(""x"",1.25)),1,0)", 1#, "FIX: non-numeric now #N/A not #VALUE!")
    gillt.Add Array("LDF_SELECT", "IF(ISNA(LDF_SELECT({1.5;2.5;3.5},""volume"",3,{10;""x"";30})),1,0)", 1#, "FIX: text in weights now #N/A not #VALUE!")
    gillt.Add Array("REINST_PREM", "IF(ISNA(REINST_PREM({8000000;""x""},5000000,5000000,500000,2)),1,0)", 1#, "FIX: no more silent text-to-zero coercion")
    gillt.Add Array("TRIANGLE", "IF(INDEX(TRIANGLE($E$2:$E$7,$F$2:$F$7,$G$2:$G$7),3,4)="""",1,0)", 1#, "FIX: dev in YEARS masks correctly (was months-only)")
    gillt.Add Array("TRIANGLE", "INDEX(TRIANGLE($E$2:$E$7,$F$2:$F$7,$G$2:$G$7),2,4)", 170#, "cumulative data: exact-lag value, cum omitted")
    gillt.Add Array("LDF_SELECT", "LDF_SELECT($J$1:$J$4)", 2.5, "FIX: empty cell in factor range skipped, not read as 0")
    gillt.Add Array("LDF_SELECT", "IF(ISNA(LDF_SELECT($J$1:$J$4,""volume"",,$K$1:$K$4)),1,0)", 1, "FIX: empty weight opposite a used factor fails loudly")
    gillt.Add Array("XINTERP", "IF(ISNA(XINTERP(0.5,$L$1:$L$5,$M$1:$M$5)),1,0)", 1, "FIX: blank table row adds no phantom (0,0) point")
    gillt.Add Array("LAYER_LOSS", "LAYER_LOSS(500000,1000000,5000000,$N$1:$N$5,$O$1:$O$5)", 307000, "FIX: blank row in ILF table ignored")
    'To edit, just add gillt.Add Array(...) following the format above
    
    Set GILL_Tests = gillt
End Function

