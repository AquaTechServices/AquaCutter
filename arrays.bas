Attribute VB_Name = "arrays"

' ShellSort an array of any type
'
' ShellSort behaves pretty well with arrays of any size, even
' if the array is already "nearly-sorted", even though in
' particular cases BubbleSort or QuickSort can be more efficient.
'
' LASTEL is the index of the last item to be sorted, and is
' useful if the array is only partially filled. This updated version accounts
' for arrays whose LBound is 0 and 1 (or whatever)
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub ShellSortS(arr() As Single, Optional lastEl As Variant,
'  Optional descending As Boolean)
'   Dim value As Single

Sub ShellSort(arr As Variant, Optional lastEl As Variant, Optional descending As Boolean)

Dim value As Variant
Dim index As Long, index2 As Long
Dim firstEl As Long
Dim distance As Long
Dim numEls As Long

' account for optional arguments
If IsMissing(lastEl) Then lastEl = UBound(arr)
firstEl = LBound(arr)
                         
numEls = lastEl - firstEl + 1
' find the best value for distance
Do
    distance = distance * 3 + 1
Loop Until distance > numEls

Do
    distance = distance \ 3
    For index = distance + firstEl To lastEl
        value = arr(index)
        index2 = index
        Do While (arr(index2 - distance) > value) Xor descending
            arr(index2) = arr(index2 - distance)
            index2 = index2 - distance
            If index2 - distance < firstEl Then Exit Do
        Loop
        arr(index2) = value
    Next
Loop Until distance = 1
End Sub

' QuickSort an array of any type
' QuickSort is especially convenient with large arrays (>1,000
' items) that contains items in random order. Its performance
' quickly degrades if the array is already almost sorted. (There are
' variations of the QuickSort algorithm that work good with
' nearly-sorted arrays, though, but this routine doesn't use them.)
'
' NUMELS is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub QuickSortS(arr() As Single, Optional numEls As Variant,
'  '     Optional descending As Boolean)
'   Dim value As Single, temp As Single

Sub QuickSort(arr As Variant, Optional numEls As Variant, Optional descending As Boolean)

Dim value As Variant, temp As Variant
Dim sp As Integer
Dim leftStk(32) As Long, rightStk(32) As Long
Dim leftNdx As Long, rightNdx As Long
Dim i As Long, j As Long

' account for optional arguments
If IsMissing(numEls) Then numEls = UBound(arr)
' init pointers
leftNdx = LBound(arr)
rightNdx = numEls
' init stack
sp = 1
leftStk(sp) = leftNdx
rightStk(sp) = rightNdx

Do
    If rightNdx > leftNdx Then
        value = arr(rightNdx)
        i = leftNdx - 1
        j = rightNdx
        ' find the pivot item
        If descending Then
            Do
                Do: i = i + 1: Loop Until arr(i) <= value
                Do: j = j - 1: Loop Until j = leftNdx Or arr(j) >= value
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            Loop Until j <= i
        Else
            Do
                Do: i = i + 1: Loop Until arr(i) >= value
                Do: j = j - 1: Loop Until j = leftNdx Or arr(j) <= value
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            Loop Until j <= i
        End If
        ' swap found items
        temp = arr(j)
        arr(j) = arr(i)
        arr(i) = arr(rightNdx)
        arr(rightNdx) = temp
        ' push on the stack the pair of pointers that differ most
        sp = sp + 1
        If (i - leftNdx) > (rightNdx - i) Then
            leftStk(sp) = leftNdx
            rightStk(sp) = i - 1
            leftNdx = i + 1
        Else
            leftStk(sp) = i + 1
            rightStk(sp) = rightNdx
            rightNdx = i - 1
        End If
    Else
        ' pop a new pair of pointers off the stacks
        leftNdx = leftStk(sp)
        rightNdx = rightStk(sp)
        sp = sp - 1
        If sp = 0 Then Exit Do
    End If
Loop
End Sub

' Indexed ShellSort of an array of any type
'
' Indexed Sorts are sort procedures that sort an index array
' instead of the main array. You can then list the items in
' sorted member by simply scanning the index, as in
'   For i = 1 To numEls: Print arr(ndx(i)): Next
'
' NUMELS is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub NdxShellSortS(arr() As Single, ndx() As Long,
'  '   Optional numEls As Variant, Optional descending As Boolean)
'   Dim value As Single

Sub NdxShellSort(arr As Variant, ndx() As Long, Optional numEls As Variant, _
Optional descending As Boolean)

Dim value As Variant
Dim index As Long, index2 As Long
Dim firstItem As Long
Dim distance As Long
Dim tempNdx As Long

' account for optional arguments
If IsMissing(numEls) Then numEls = UBound(arr)
firstItem = LBound(arr)
' init index array if necessary
If ndx(firstItem) = 0 And ndx(UBound(ndx)) = 0 Then
    For index = firstItem To UBound(ndx)
        ndx(index) = index
    Next
End If
                         
' find the best value for distance
Do
    distance = distance * 3 + 1
Loop Until distance > numEls
                         
Do
    distance = distance \ 3
    For index = distance + 1 To numEls
        tempNdx = ndx(index)
        value = arr(tempNdx)
        index2 = index
        Do While (arr(ndx(index2 - distance)) > value) Xor descending
            ndx(index2) = ndx(index2 - distance)
            index2 = index2 - distance
            If index2 <= distance Then Exit Do
        Loop
        ndx(index2) = tempNdx
    Next
Loop Until distance = 1
End Sub

' Returns True if an array contains duplicate values
' it works with arrays of any type

Function HasDuplicateValues(arr As Variant) As Boolean
Dim col As Collection, index As Long
Set col = New Collection
                         
' assume that the array contains duplicates
HasDuplicateValues = True
                         
On Error GoTo FoundDuplicates
For index = LBound(arr) To UBound(arr)
    ' build the key using the array element
    ' an error occurs if the key already exists
    col.Add 0, CStr(arr(index))
Next
' if control comes here, the array doesn't contain
' any duplicate values, so we can return zero
HasDuplicateValues = False
                         
FoundDuplicates:

End Function

' Filter out duplicate values in an array and compact
' the array by moving items to "fill the gaps".
' Returns the number of duplicate values
'
' it works with arrays of any type, except objects
'
' The array is not REDIMed, but you can do it easily using
' the following code:
'     a() is a string array
'     dups = FilterDuplicates(a())
'     If dups Then
'         ReDim Preserve a(LBound(a) To UBound(a) - dups) As String
'     End If

Function FilterDuplicates(arr As Variant) As Long
Dim col As Collection, index As Long, dups As Long
Set col = New Collection
                         
On Error Resume Next
                         
For index = LBound(arr) To UBound(arr)
' build the key using the array element
' an error occurs if the key already exists
col.Add 0, CStr(arr(index))
If Err Then
    ' we've found a duplicate
    arr(index) = Empty
    dups = dups + 1
    Err.Clear
ElseIf dups Then
    ' if we've found one or more duplicates so far
    ' we need to move elements towards lower indices
    arr(index - dups) = arr(index)
    arr(index) = Empty
End If
Next
                         
' return the number of duplicates
FilterDuplicates = dups
                         
End Function

' Bubble Sort an array of any type
' BubbleSort is especially convenient with small arrays (1,000
' items or fewer) or with arrays that are already almost sorted
'
' NUMELS is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub BubbleSortS(arr() As Single, Optional numEls As Variant,
'  '     Optional descending As Boolean)
'   Dim value As Single

Sub BubbleSort(arr As Variant, Optional numEls As Variant, Optional descending As Boolean)

Dim value As Variant
Dim index As Long
Dim firstItem As Long
Dim indexLimit As Long, lastSwap As Long

' account for optional arguments
If IsMissing(numEls) Then numEls = UBound(arr)
firstItem = LBound(arr)
lastSwap = numEls

Do
    indexLimit = lastSwap - 1
    lastSwap = 0
    For index = firstItem To indexLimit
        value = arr(index)
        If (value > arr(index + 1)) Xor descending Then
            ' if the items are not in order, swap them
            arr(index) = arr(index + 1)
            arr(index + 1) = value
            lastSwap = index
        End If
    Next
Loop While lastSwap
End Sub

' Binary search in an array of any type
' Returns the index of the matching item, or -1 if the search fails
'
' The arrays *must* be sorted, in ascending or descending
' order (the routines finds out the sort direction).
' LASTEL is the index of the last item to be searched, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, including objects if your are searching
' for their default property, and excluding UDTs and fixed-length strings.
' String are compared in case-sensitive mode.
'
' You can write faster procedures if you modify the first line
' to account for a specific data type, eg.
'   Function BinarySearchL (arr() As Long, search As Long,
'  Optional lastEl As Variant) As Long

Function BinarySearch(arr As Variant, search As Variant, Optional lastEl As Variant) As Long
Dim index As Long
Dim first As Long
Dim last As Long
Dim middle As Long
Dim inverseOrder As Boolean
                         
' account for optional arguments
If IsMissing(lastEl) Then lastEl = UBound(arr)
' deduct direction of sorting
inverseOrder = (arr(first) > arr(last))
                         
first = LBound(arr)
last = lastEl
' assume searches failed
BinarySearch = first - 1
                         
Do
    middle = (first + last) \ 2
    If arr(middle) = search Then
        BinarySearch = middle
        Exit Do
    ElseIf ((arr(middle) < search) Xor inverseOrder) Then
        first = middle + 1
    Else
        last = middle - 1
    End If
Loop Until first > last
End Function
                     
' Return the sum of the values in an array of any type
' (for string arrays, it concatenates all its elements)
'
' FIRST and LAST indicate which portion of the array
' should be considered; they default to the first
' and last element, respectively

Function ArraySum(arr As Variant, Optional first As Variant, Optional last As Variant) As Variant
Dim index As Long

If IsMissing(first) Then first = LBound(arr)
If IsMissing(last) Then last = UBound(arr)

For index = first To last
    ArraySum = ArraySum + arr(index)
Next
End Function
                     
' The standard deviation of an array of any type
'
' if the second argument is True or omitted,
' it evaluates the standard deviation of a sample,
' if it is False it evaluates the standard deviation of a population
'
' if the third argument is True or omitted, Empty values aren't accounted for

Function ArrayStdDev(arr As Variant, Optional SampleStdDev As Boolean = True, Optional IgnoreEmpty As Boolean = True) As Double
Dim sum As Double
Dim sumSquare As Double
Dim value As Double
Dim count As Long
Dim index As Long

' evaluate sum of values
' if arr isn't an array, the following statement raises an error
For index = LBound(arr) To UBound(arr)
    value = arr(index)
    ' skip over non-numeric values
    If IsNumeric(value) Then
        ' skip over empty values, if requested
        If Not (IgnoreEmpty And IsEmpty(value)) Then
            ' add to the running total
            count = count + 1
            sum = sum + value
            sumSquare = sumSquare + value * value
        End If
    End If
Next

' evaluate the result
' use (Count-1) if evaluating the standard deviation of a sample
If SampleStdDev Then
    ArrayStdDev = Sqr((sumSquare - (sum * sum / count)) / (count - 1))
Else
    ArrayStdDev = Sqr((sumSquare - (sum * sum / count)) / count)
End If

End Function
                     
' Shuffle the elements of an array of any type
' (it doesn't work with arrays of objects or UDT)

Sub ArrayShuffle(arr As Variant)
    Dim index As Long
    Dim newIndex As Long
    Dim firstIndex As Long
    Dim itemCount As Long
    Dim tmpValue As Variant
                         
    firstIndex = LBound(arr)
    itemCount = UBound(arr) - LBound(arr) + 1
                         
    For index = UBound(arr) To LBound(arr) + 1 Step -1
        ' evaluate a random index from LBound to INDEX
        newIndex = firstIndex + Int(Rnd * itemCount)
        ' swap the two items
        tmpValue = arr(index)
        arr(index) = arr(newIndex)
        arr(newIndex) = tmpValue
        ' prepare for next iteration
        itemCount = itemCount - 1
    Next

End Sub

' The average of an array of any type
'
' FIRST and LAST indicate which portion of the array
' should be considered; they default to the first
' and last element, respectively
' if IGNOREEMPTY argument is True or omitted,
' Empty values aren't accounted for

Function ArrayAvg(arr As Variant, Optional first As Variant, Optional last As Variant, Optional IgnoreEmpty As Boolean = True) As Variant
    Dim index As Long
    Dim sum As Variant
    Dim count As Long

If IsMissing(first) Then first = LBound(arr)
If IsMissing(last) Then last = UBound(arr)
                         
' if arr isn't an array, the following statement raises an error
For index = first To last
    If IgnoreEmpty = False Or Not IsEmpty(arr(index)) Then
        sum = sum + arr(index)
        count = count + 1
    End If
Next
                         
' return the average
ArrayAvg = sum / count

End Function

' Cotangent of an angle

Function Cot(radians As Double) As Double
    Cot = 1 / Tan(radians)
End Function

' Secant of an angle

Function Sec(radians As Double) As Double
    Sec = 1 / Cos(radians)
End Function



