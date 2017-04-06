 
Public Sub TestJSON13()

    Dim parser As New JSON13
    Dim jsonText As String
    jsonText = "{ ""t1"":""testvalue1"", ""escaped"":""x\tx\""x\u20AC3,000.00"", ""b"":999, ""c"":[ 1,2,3,4,5,], ""d"":{ ""d1"":""D1"", ""d2"":""D2""}, ""e"":{""d1"":[ 1,2,3,4,5,],""d2"":[ 1,2,3,4,5,], }, ""array_of_objects"":[ {}, [], {""x"":1}], ""success"":""True"", ""last"":99.99 }"

    Dim doc As Variant
    Set doc = parser.parse(jsonText)

    ' access the top level properties
    Debug.Print "TestJSON: property 't1' is " & doc.getString("t1")
    Debug.Print "TestJSON: property 'escaped' is " & doc.getString("escaped") & " is " & (doc.getString("escaped") = "testvalue1")
    Debug.Print "TestJSON: property 'success' is " & doc.getString("success")

    ' everything is wrapped in a JSON13 object
    ' example of how to access an array
    Dim prop_c As JSON13
    Set prop_c = doc.getObj("c")
    
    Dim myArray As Variant
    Set myArray = prop_c.value
    
    Dim i As Variant
    For Each i In myArray
        ' every array is a collection, and every element of the array is another JSON13 object wrapping a value
        Debug.Print "TestJSON: property 'c' element = " & i.value
    Next i
    
    Call doc.dump
    Debug.Print "TestJSON: doc.toString(): " & doc.toString()
End Sub

