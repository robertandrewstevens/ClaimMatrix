'*************************************************************************
'  Program:  ClaimMatrix                                                 *
'   Author:  Robert Andrew Stevens                                       *
'     Date:  10/25/04                                                    *
'  Purpose:  Create a Production vs. Time-In-Service Matrix              *
'            for Weibull++                                               *
'*************************************************************************
 
Const MaxData As Integer = 3100 ' Maximum number of records
Const MaxVec As Integer = 20 ' Maximum number of elements in vectors
 
Type Warr_Struct       ' Warranty Data Structure
    Claim As String    ' Claim No (Col = 1)
    Part As String     ' Part Number (Col = 2)
    PIN As String      ' PIN Number (Col = 7)
    BldDate As String  ' Build Date (Col = 11)
    FailDate As String ' Failure Date (Col = 13)
End Type
Dim Warranty(1 To MaxData) As Warr_Struct ' Warranty data
 
Type Mach_Struct                    ' Machine Data Structure
    PIN As String                   ' PIN Number
    Claims As Integer               ' Numer of Claims
    BldDate(1 To MaxVec) As String  ' Build Date - can change if replacement part fails
    Part(1 To MaxVec) As String     ' Part Number
    FailDate(1 To MaxVec) As String ' Failure Date
    Count(1 To MaxVec) As Boolean   ' True = Yes, include failure; False = No, do not include failure
End Type
Dim Mach(1 To MaxData) As Mach_Struct ' Machine data
Dim Machines(1 To MaxData) As String  ' Vector of unique machines
Dim Num_Claims As Integer             ' Number of Claim data records
Dim Num_Mach As Integer               ' Number of Machine data records
 
Dim BldDates(1 To MaxData) As String  ' Vector of unique Build Dates
Dim BldYears(1 To MaxData) As String  ' Vector of unique Build Years
Dim BldMonths(1 To MaxData) As String ' Vector of unique Build Months
Dim BldYYYYMM(1 To MaxData) As String ' Vector of unique Build Year-Month combination
Dim BldCount(1 To MaxData) As Integer ' Number of machines built in Year-Month
 
Dim Num_BldDates As Integer  ' Number of Build Date data records
Dim Num_BldYears As Integer  ' Number of Build Year data records
Dim Num_BldMonths As Integer ' Number of Build Month data records
Dim Num_BldYYYYMM As Integer ' Number of Build Month data records
 
Dim FailDates(1 To MaxData) As String  ' Vector of unique Failure Dates
Dim FailYears(1 To MaxData) As String  ' Vector of unique Failure Years
Dim FailMonths(1 To MaxData) As String ' Vector of unique Failure Months
Dim FailYYYYMM(1 To MaxData) As String ' Vector of unique Build Year-Month combination
Dim FailCount(1 To MaxData) As Integer ' Number of failures in Year-Month
 
Dim Num_FailDates As Integer  ' Number of Build Date data records
Dim Num_FailYears As Integer  ' Number of Build Year data records
Dim Num_FailMonths As Integer ' Number of Build Month data records
Dim Num_FailYYYYMM As Integer ' Number of Build Month data records
 
Dim FailMatrix(1 To MaxData, 1 To MaxData) As Integer ' Matrix of failures by build and failure months
 
Sub Main()
 
    Read_Claims  ' Read warranty claims
    Create_Mach  ' Create machine data structure
    Build_Months ' Determine months of build
    Fail_Months  ' Determine months of failures
    Count_Build  ' Count number of machines in build month
    Count_Fail   ' Count number of failures in failure month
    Filter       ' Filter out parts that are not to be included
    Make_Matrix  ' Create matrix of failure by build and failure month
 
End Sub
 
'**************************************************************************
' function: Read_Claims                                                   *
'  purpose: read Claim/Cost data spreadsheet and store in array           *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Read_Claims()
 
    Set Claims = ThisWorkbook.Sheets("Claims")
    Set Claims1 = ThisWorkbook.Sheets("Claims1")
    Dim i As Integer ' Loop counter
 
    Claims1.Select
    Cells.Select
    Selection.ClearContents
 
    Num_Claims = Application.CountA(Claims.Range("A:A")) - 1 ' Subtract 1 for Header
 
    For i = 1 To Num_Claims
 
        Warranty(i).Claim = Claims.Cells(i + 1, 1) ' Add 1 for Header
        'Claims1.Cells(i, 1) = Warranty(i).Claim ' Echo input
 
        Warranty(i).Part = Claims.Cells(i + 1, 2) ' Add 1 for Header
        'Claims1.Cells(i, 2) = Warranty(i).Part ' Echo input
 
        Warranty(i).PIN = Claims.Cells(i + 1, 7) ' Add 1 for Header
        'Claims1.Cells(i, 3) = Warranty(i).PIN ' Echo input
 
        Warranty(i).BldDate = Format(Claims.Cells(i + 1, 11), "yyyy/mm/dd") ' Add 1 for Header
        'Claims1.Cells(i, 4) = Warranty(i).BldDate ' Echo input
 
        Warranty(i).FailDate = Format(Claims.Cells(i + 1, 13), "yyyy/mm/dd") ' Add 1 for Header
        'Claims1.Cells(i, 5) = Warranty(i).FailDate ' Echo input
 
    Next i
 
    'MsgBox ("Number of records read = " & Num_Claims)
End Sub
 
'**************************************************************************
' function: Create_Mach                                                   *
'  purpose: build the Machine data structure                              *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Create_Mach()
 
    Dim i, j, k As Integer  ' Loop counters
    Set MachData = ThisWorkbook.Sheets("MachData")
 
    MachData.Select
    Cells.Select
    Selection.ClearContents
 
'   Vector size
    Num_Mach = Num_Claims
    Num_FailDates = Num_Claims
 
'   Get values from Warranty database
    For i = 1 To Num_Claims
        Machines(i) = Warranty(i).PIN
        FailDates(i) = Warranty(i).FailDate
    Next i
 
'   Create vector of unique values
    Call Uniq_List(Machines, Num_Mach, "Machines")
    Call Uniq_List(FailDates, Num_FailDates, "FailDates")
 
'   Fill Machine data structure
    For i = 1 To Num_Mach
        Mach(i).PIN = Machines(i)
        k = 0
        For j = 1 To Num_Claims
            If Mach(i).PIN = Warranty(j).PIN Then
                k = k + 1
                Mach(i).BldDate(k) = Warranty(j).BldDate ' Initial value
                Mach(i).Part(k) = Warranty(j).Part
                Mach(i).FailDate(k) = Warranty(j).FailDate
            End If
        Next j
        Mach(i).Claims = k
    Next i
 
'   Adjust BldDate if repeat failure(s)
    For i = 1 To Num_Mach
        For j = 1 To Mach(i).Claims - 1
            For k = j + 1 To Mach(i).Claims
                If Mach(i).Part(j) = Mach(i).Part(k) Then
                    Mach(i).BldDate(k) = Mach(i).FailDate(j) ' Reset BldDate for repeat failure to the previous FailDate
                End If
            Next k
        Next j
    Next i
 
'   Write data to spreadsheet to check
    k = 1
    For i = 1 To Num_Mach
        For j = 1 To Mach(i).Claims
            MachData.Cells(k, 1) = Mach(i).PIN
            MachData.Cells(k, 2) = Mach(i).Claims
            MachData.Cells(k, 3) = Mach(i).Part(j)
            MachData.Cells(k, 4) = Mach(i).BldDate(j)
            MachData.Cells(k, 5) = Mach(i).FailDate(j)
            k = k + 1
        Next j
    Next i
 
'   Get build dates from Machine data
    k = 0
    For i = 1 To Num_Mach
        k = k + 1
        BldDates(k) = Mach(i).BldDate(1) ' Start with orginal build dates
        For j = 1 To Mach(i).Claims
            k = k + 1
            BldDates(k) = Mach(i).FailDate(j) ' Add failure dates for replacement parts
        Next j
    Next i
    Num_BldDates = k
 
'   Create vector of unique values
    Call Uniq_List(BldDates, Num_BldDates, "BldDates")
 
End Sub
 
'**************************************************************************
' function: Build_Months                                                  *
'  purpose: Determine build months                                        *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Build_Months()
    Set Build = ThisWorkbook.Sheets("BldDates")
    Dim i, j As Integer ' Loop counter
 
    Build.Select
    Cells.Select
    Selection.ClearContents
 
'   Initialize vector size
    Num_BldYears = Num_BldDates
    Num_BldMonths = Num_BldDates
    Num_BldYYYYMM = Num_BldDates
 
'   Get values from Warranty database and increment counter variables
    For i = 1 To Num_BldDates
        Build.Cells(i, 1) = BldDates(i) ' Echo input
        ' Determine Year
        BldYears(i) = Mid(Format(BldDates(i), "yyyy/mm/dd"), 1, 4)
        Build.Cells(i, 2) = BldYears(i)
        ' Determine Month
        BldMonths(i) = Mid(Format(BldDates(i), "yyyy/mm/dd"), 6, 2)
        Build.Cells(i, 3) = BldMonths(i)
        '  Determine Day
        Build.Cells(i, 4) = Mid(Format(BldDates(i), "yyyy/mm/dd"), 9, 4)
        '  Determine Year-Month combination
        BldYYYYMM(i) = BldYears(i) & BldMonths(i)
    Next i
 
'   Create vector of unique values
    Call Uniq_List(BldYears, Num_BldYears, "BldYears")
    Call Uniq_List(BldMonths, Num_BldMonths, "BldMonths")
    Call Uniq_List(BldYYYYMM, Num_BldYYYYMM, "BldYYYYMM")
 
End Sub
 
'**************************************************************************
' function: Fail_Months                                                   *
'  purpose: Determine months that failures occurred                       *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Fail_Months()
    Set Output = ThisWorkbook.Sheets("FailDates")
    Dim i, j As Integer ' Loop counter
 
    Output.Select
    Cells.Select
    Selection.ClearContents
 
'   Initialize vector sizes
    Num_FailYears = Num_FailDates
    Num_FailMonths = Num_FailDates
    Num_FailYYYYMM = Num_FailDates
 
'   Get values from Warranty database
    For i = 1 To Num_FailDates
        Output.Cells(i, 1) = FailDates(i) ' Echo input
        ' Determine Year
        FailYears(i) = Mid(Format(FailDates(i), "yyyy/mm/dd"), 1, 4)
        Output.Cells(i, 2) = FailYears(i)
        ' Determine Month
        FailMonths(i) = Mid(Format(FailDates(i), "yyyy/mm/dd"), 6, 2)
        Output.Cells(i, 3) = FailMonths(i)
        '  Determine Day
        Output.Cells(i, 4) = Mid(Format(FailDates(i), "yyyy/mm/dd"), 9, 4)
        '  Determine Year-Month combination
        FailYYYYMM(i) = FailYears(i) & FailMonths(i)
    Next i
 
'   Create vector of unique values
    Call Uniq_List(FailYears, Num_FailYears, "FailYears")
    Call Uniq_List(FailMonths, Num_FailMonths, "FailMonths")
    Call Uniq_List(FailYYYYMM, Num_FailYYYYMM, "FailYYYYMM")
 
End Sub
 
'**************************************************************************
' function: Count_Build                                                   *
'  purpose: Determine number of machines built in month (from Claim data) *
'           Note:  can overwrite this with actual data later              *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Count_Build()
 
    Dim i, j, k As Integer  ' Loop counters
    Set Output = ThisWorkbook.Sheets("BldYYYYMM")
    Dim Build_Year As String
    Dim Build_Month As String
    Dim Mach_Year As String
    Dim Mach_Month As String
 
'    For i = 1 To Num_BldYYYYMM
'        BldCount(i) = 0
'        Build_Year = Mid(BldYYYYMM(i), 1, 4)
'        Build_Month = Mid(BldYYYYMM(i), 5, 2)
'        For j = 1 To Num_Mach
'            For k = 1 To Mach(j).Claims
'                Mach_Year = Mid(Format(Mach(j).BldDate(k), "yyyy/mm/dd"), 1, 4)
'                Mach_Month = Mid(Format(Mach(j).BldDate(k), "yyyy/mm/dd"), 6, 2)
'                If Build_Year = Mach_Year And Build_Month = Mach_Month Then
'                    BldCount(i) = BldCount(i) + 1
'                End If
'            Next k
'        Next j
'    Next i
 
    For i = 1 To Num_BldYYYYMM
        BldCount(i) = 0
        Build_Year = Mid(BldYYYYMM(i), 1, 4)
        Build_Month = Mid(BldYYYYMM(i), 5, 2)
        For j = 1 To Num_Mach
            Mach_Year = Mid(Format(Mach(j).BldDate(1), "yyyy/mm/dd"), 1, 4)
            Mach_Month = Mid(Format(Mach(j).BldDate(1), "yyyy/mm/dd"), 6, 2)
            If Build_Year = Mach_Year And Build_Month = Mach_Month Then
                BldCount(i) = BldCount(i) + 1 ' Initial value - adjust later
            End If
        Next j
    Next i
 
'   Write data to spreadsheet to check
    For i = 1 To Num_BldYYYYMM
        Output.Cells(i, 2) = BldCount(i)
    Next i
 
End Sub
 
'**************************************************************************
' function: Count_Fail                                                    *
'  purpose: Determine number of falures in a month                        *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Count_Fail()
 
    Dim i, j, k As Integer  ' Loop counters
    Set Output = ThisWorkbook.Sheets("FailYYYYMM")
    Dim Fail_Year As String
    Dim Fail_Month As String
    Dim Mach_Year As String
    Dim Mach_Month As String
 
    For i = 1 To Num_FailYYYYMM
        FailCount(i) = 0
        Fail_Year = Mid(FailYYYYMM(i), 1, 4)
        Fail_Month = Mid(FailYYYYMM(i), 5, 2)
        For j = 1 To Num_Mach
            For k = 1 To Mach(j).Claims
                Mach_Year = Mid(Format(Mach(j).FailDate(k), "yyyy/mm/dd"), 1, 4)
                Mach_Month = Mid(Format(Mach(j).FailDate(k), "yyyy/mm/dd"), 6, 2)
                If Fail_Year = Mach_Year And Fail_Month = Mach_Month Then
                    FailCount(i) = FailCount(i) + 1
                End If
            Next k
        Next j
    Next i
 
'   Write data to spreadsheet to check
    For i = 1 To Num_FailYYYYMM
        Output.Cells(i, 2) = FailCount(i)
    Next i
 
End Sub
 
'**************************************************************************
' function: Filter                                                        *
'  purpose: Read in list of parts and set "count" for failure             *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Filter()
 
    Set Parts = ThisWorkbook.Sheets("Parts")
 
    Dim i, j, k As Integer ' Loop counters
    Dim Num_Flt As Integer ' Number of parts in list
    Dim Parts_List(1 To MaxData) As String
 
'   Read parts list
    Num_Parts = Application.CountA(Parts.Range("A:A"))
    For i = 1 To Num_Parts
        Parts_List(i) = Parts.Cells(i, 1)
    Next i
 
'   Determine whether Claim part on Machine matches part on list
    For i = 1 To Num_Parts
        For j = 1 To Num_Mach
            For k = 1 To Mach(j).Claims
                Mach(j).Count(k) = False ' Initialize
                If Mach(j).Part(k) = Parts_List(i) Then
                    Mach(j).Count(k) = True
                End If
            Next k
        Next j
    Next i
 
End Sub
 
'**************************************************************************
' function: Make_Matrix                                                   *
'  purpose: Create a matrix of failures by month of build and failure     *
'    input:                                                               *
'   output:                                                               *
'**************************************************************************
Sub Make_Matrix()
 
    Dim i, j, k, l, m As Integer  ' Loop counters
    Set Output = ThisWorkbook.Sheets("Matrix")
    Dim Bld_Year As String
    Dim Bld_Month As String
    Dim Bld_Year2 As String
    Dim Bld_Month2 As String
    Dim Fail_Year As String
    Dim Fail_Month As String
    Dim Mach_Bld_Year As String
    Dim Mach_Bld_Month As String
    Dim Mach_Fail_Year As String
    Dim Mach_Fail_Month As String
 
    For i = 1 To Num_BldYYYYMM
        Bld_Year = Mid(BldYYYYMM(i), 1, 4)
        Bld_Month = Mid(BldYYYYMM(i), 5, 2)
        For j = 1 To Num_FailYYYYMM
            FailMatrix(i, j) = 0
            Fail_Year = Mid(FailYYYYMM(j), 1, 4)
            Fail_Month = Mid(FailYYYYMM(j), 5, 2)
            For k = 1 To Num_Mach
                For l = 1 To Mach(k).Claims
                    Mach_Bld_Year = Mid(Format(Mach(k).BldDate(l), "yyyy/mm/dd"), 1, 4)
                    Mach_Bld_Month = Mid(Format(Mach(k).BldDate(l), "yyyy/mm/dd"), 6, 2)
                    If Mach(k).Count(l) Then ' Part included on list
                        Mach_Fail_Year = Mid(Format(Mach(k).FailDate(l), "yyyy/mm/dd"), 1, 4)
                        Mach_Fail_Month = Mid(Format(Mach(k).FailDate(l), "yyyy/mm/dd"), 6, 2)
                        If Bld_Year = Mach_Bld_Year And Bld_Month = Mach_Bld_Month And Fail_Year = Mach_Fail_Year And Fail_Month = Mach_Fail_Month Then
                            FailMatrix(i, j) = FailMatrix(i, j) + 1
                            For m = i To Num_BldYYYYMM
                                Bld_Year2 = Mid(BldYYYYMM(m), 1, 4)
                                Bld_Month2 = Mid(BldYYYYMM(m), 5, 2)
                                If Bld_Year2 = Fail_Year And Bld_Month2 = Fail_Month Then
                                    BldCount(m) = BldCount(m) + 1 ' increase build count in that month
                                End If
                            Next m
                        End If
                    End If
                Next l
            Next k
        Next j
    Next i
 
'   Write data to spreadsheet to check
    Output.Select
    Cells.Select
    Selection.ClearContents
 
    For j = 1 To Num_FailYYYYMM
        Output.Cells(1, j + 2) = FailYYYYMM(j) ' Column labels
    Next j
 
    For i = 1 To Num_BldYYYYMM
        Output.Cells(i + 1, 1) = BldYYYYMM(i) ' Row labels
        Output.Cells(i + 1, 2) = BldCount(i) ' Number of machines built
        For j = 1 To Num_FailYYYYMM
            Output.Cells(i + 1, j + 2) = FailMatrix(i, j)
        Next j
    Next i
 
End Sub

 
'**************************************************************************
' function: Uniq_List                                                     *
'  purpose: Create list of unique values                                  *
'    input: List name, number of elements, and location to print          *
'   output:                                                               *
'**************************************************************************
Sub Uniq_List(List() As String, NList As Integer, Location As String)
 
    Dim i As Integer
 
    Call Sort_List(List, NList)
    Call Del_Rep(List, NList)
    Call Print_List(List, NList, Location)
End Sub
 
'*************************************************************************
'  Function:  sort_list                                                  *
'   Purpose:  sort list of strings                                       *
'    Inputs:  list name and number of elements                           *
'    Return:                                                             *
'*************************************************************************
Sub Sort_List(List() As String, NList As Integer)
 
    Dim i, j As Integer ' Loop counters
    Dim Tmp_Str As String
 
    For i = 1 To NList - 1
        For j = i + 1 To NList
            If List(i) > List(j) Then
                Tmp_Str = List(i)
                List(i) = List(j)
                List(j) = Tmp_Str
            End If
        Next j
    Next i
End Sub
 
'*************************************************************************
'  Function:  del_rep                                                    *
'   Purpose:  delete repeats in a list of strings                        *
'    Inputs:  list name and number of elements                           *
'    Return:                                                             *
'*************************************************************************
Sub Del_Rep(List() As String, NList As Integer)
 
    Dim i, j As Integer ' Loop counters
 
    i = 2
    Do While i <= NList
        If List(i) = List(i - 1) Then
            If Not i = NList Then
                For j = i To NList
                    List(j) = List(j + 1)
                Next j
            End If
            NList = NList - 1
        Else: i = i + 1
        End If
    Loop
End Sub
 
'*************************************************************************
'  Function:  Print_list                                                 *
'   Purpose:  Print list to a sheet                                      *
'    Inputs:  list name, number of elements, and location to print       *
'    Return:                                                             *
'*************************************************************************
Sub Print_List(List() As String, NList As Integer, Location As String)
 
    Dim i As Integer ' Loop counter
 
    Sheets(Location).Select
    Cells.Select
    Selection.ClearContents
    For i = 1 To NList
        Sheets(Location).Cells(i, 1) = List(i)
    Next i
End Sub


Sub HelpButton()
    Dim HelpDlg As DialogSheet
    Set HelpDlg = ThisWorkbook.DialogSheets("HelpDlg")
 
    HelpDlg.Show
End Sub
