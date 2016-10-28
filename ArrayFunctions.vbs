Option Explicit

Function getDimension(var As Variant) As Integer
On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    i = 0
    Do While True:
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getDimension = i - 1
End Function

Function joinArray(arr1() As String, arr2() As String) As String()
    Dim result() As String
    Dim tmp1(), tmp2() As String
    If getDimension(arr1) = 1 And getDimension(arr2) = 1 Then
        joinArray = joinArray1(arr1, arr2)
    Else
        If getDimension(arr1) >= getDimension(arr2) Then
            While getDimension(arr2) > 1
                arr1 = joinArray2(arr1, splitArray1(arr2))
                arr2 = splitArray2(arr2)
            Wend
            arr2 = RemoveDuplicated(arr2)
            joinArray = joinArray2(arr1, arr2)
        Else
            result = joinArray(arr2, arr1)
            joinArray = result
        End If
    End If
End Function

Function joinArray1(arr1() As String, arr2() As String) As String()
    Dim result() As String
    Dim len1, len2, len3, k, i, j As Integer
    len1 = UBound(arr1)
    len2 = UBound(arr2)
    len3 = len1 * len2
    ReDim result(len3, 2)
    k = 1
    For i = 1 To len1
        For j = 1 To len2
            result(k, 1) = arr1(i)
            result(k, 2) = arr2(j)
            k = k + 1
        Next
    Next
    joinArray1 = result
End Function

Function joinArray2(arr1() As String, arr2() As String) As String()
    Dim result() As String
    Dim len11, len12, len2, i, j, k, q As Integer
    len11 = UBound(arr1, 1)
    len12 = UBound(arr1, 2)
    len2 = UBound(arr2)
    ReDim result(len11 * len2, len12 + 1)
    k = 1
    For i = 1 To len11
        For j = 1 To len2
            For q = 1 To len12
                result(k, q) = arr1(i, q)
            Next
            result(k, len12 + 1) = arr2(j)
            k = k + 1
        Next
    Next
    joinArray2 = result
End Function

Function splitArray1(arr() As String) As String()
    Dim result() As String
    Dim myLen, i As Integer
    myLen = UBound(arr)
    ReDim result(myLen)
    For i = 1 To myLen
        result(i) = arr(i, 1)
    Next
    splitArray1 = RemoveDuplicated(result)
End Function

Function splitArray2(arr() As String) As String()
    Dim result() As String
    Dim myLen As Integer
    Dim myLen1, myLen2, i, j As Integer
    
    myLen1 = UBound(arr, 1)
    myLen2 = UBound(arr, 2)
    If myLen2 > 2 Then
        ReDim result(myLen1, myLen2 - 1)
        For i = 1 To myLen1
            For j = 2 To myLen2
                result(i, j - 1) = arr(i, j)
            Next
        Next
    Else
        ReDim result(myLen1)
        For i = i To myLen1
            result(i) = arr(i, 2)
        Next
    End If
    splitArray2 = result
End Function

Sub printArray(arr() As String)
    Dim i, j As Integer
    Dim result As String
    
    If getDimension(arr) = 1 Then
        For i = 1 To UBound(arr)
            Debug.Print arr(i)
        Next
    ElseIf getDimension(arr) = 2 Then
        For i = 1 To UBound(arr, 1)
            result = ""
            For j = 1 To UBound(arr, 2)
                result = result & CStr(arr(i, j)) & ","
            Next
            result = Left(result, Len(result) - 1)
            Debug.Print result
        Next
    End If
End Sub

Function RemoveDuplicated(Array_1() As String) As String()
Dim Array_2() As String
ReDim Array_2(1)
Dim x, i As Integer

Array_2(1) = Array_1(1)
x = 2
For i = 2 To UBound(Array_1)
    If UBound(Filter(Array_2, Array_1(i))) = -1 Then
        ReDim Preserve Array_2(x)
        Array_2(x) = Array_1(i)
        x = x + 1
    End If
Next
RemoveDuplicated = Array_2
End Function

Function RangeToArray(rng As Range) As String()
    Dim result() As String
    Dim nbCol, i, j As Integer
    If rng.Columns.Count = 1 Then
        ReDim result(rng.Rows.Count)
        For i = 1 To UBound(result)
            result(i) = rng.Cells(i, 1).Value
        Next
    Else
        ReDim result(rng.Rows.Count, rng.Columns.Count)
        For i = 1 To UBound(result, 1)
            For j = 1 To UBound(result, 2)
                result(i, j) = rng.Cells(i, j).Value
            Next
        Next
    End If
    RangeToArray = result
End Function

Sub pasteArray(arr() As String, dst As Range)
    Dim i, j As Integer
    
    If getDimension(arr) = 1 Then
        For i = 1 To UBound(arr)
            dst.Cells(i, 1).Value = arr(i)
        Next
    Else
        For i = 1 To UBound(arr, 1)
            For j = 1 To UBound(arr, 2)
                dst.Cells(i, j).Value = arr(i, j)
            Next
        Next
    End If
End Sub

Sub joinRange(rng1 As Range, rng2 As Range, dst As Range)
    Dim arr() As String, arr1() As String, arr2() As String
    
    arr1 = RangeToArray(rng1)
    arr2 = RangeToArray(rng2)
    arr = joinArray(arr1, arr2)
    pasteArray arr, dst
End Sub

Function filterArray(tabIni() As String, pattern As String, col As Integer) As String()
    Dim result() As String
    Dim i As Integer
    Dim cnt As Integer
    
    cnt = 0
    For i = 1 To UBound(tabIni)
        If tabIni(i, 1) = pattern Then
            cnt = cnt + 1
        End If
    Next
    
    ReDim result(cnt, 3)
    cnt = 1
    For i = 1 To UBound(tabIni)
        If tabIni(i, col) = pattern Then
            result(cnt, 1) = pattern
            result(cnt, 2) = tabIni(i, 2)
            result(cnt, 3) = tabIni(i, 3)
            cnt = cnt + 1
        End If
    Next
    
    filterArray = result
    
End Function
