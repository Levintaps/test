Sub GetHPProductName()
    Dim serialNumber As String
    Dim productName As String
    Dim currentRow As Long
    Dim lastRow As Long
    Dim http As Object
    Dim html As Object
    Dim url As String
    Dim response As String
    
    ' Find the last row with data in column B
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    ' Loop through all rows with serial numbers
    For currentRow = 1 To lastRow
        serialNumber = Trim(Cells(currentRow, "B").Value)
        
        ' Skip if serial number is empty or product name already exists
        If serialNumber <> "" And Cells(currentRow, "A").Value = "" Then
            
            ' Create HTTP request object
            Set http = CreateObject("MSXML2.XMLHTTP.6.0")
            
            ' HP warranty check URL with serial number
            url = "https://support.hp.com/ph-en/check-warranty"
            
            ' Set up the request
            http.Open "POST", url, False
            http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            
            ' Send request with serial number
            http.Send "serialNumber=" & serialNumber
            
            ' Get response
            response = http.responseText
            
            ' Parse the response to extract product name
            productName = ExtractProductName(response)
            
            ' If product name found, put it in column A
            If productName <> "" Then
                Cells(currentRow, "A").Value = productName
            Else
                Cells(currentRow, "A").Value = "Product not found"
            End If
            
            ' Add small delay to avoid overwhelming the server
            Application.Wait (Now + TimeValue("0:00:01"))
            
            ' Clean up
            Set http = Nothing
            
        End If
    Next currentRow
    
    MsgBox "Product lookup completed!"
End Sub

Function ExtractProductName(htmlContent As String) As String
    Dim productName As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Look for common patterns in HP warranty pages
    ' These patterns may need adjustment based on actual HP page structure
    
    ' Pattern 1: Look for product name in title or heading
    startPos = InStr(htmlContent, "product-name")
    If startPos > 0 Then
        startPos = InStr(startPos, htmlContent, ">") + 1
        endPos = InStr(startPos, htmlContent, "<")
        If endPos > startPos Then
            productName = Mid(htmlContent, startPos, endPos - startPos)
            productName = Trim(Replace(Replace(productName, vbCrLf, ""), vbLf, ""))
            If productName <> "" Then
                ExtractProductName = productName
                Exit Function
            End If
        End If
    End If
    
    ' Pattern 2: Look for warranty info section
    startPos = InStr(htmlContent, "warranty-product-name")
    If startPos > 0 Then
        startPos = InStr(startPos, htmlContent, ">") + 1
        endPos = InStr(startPos, htmlContent, "<")
        If endPos > startPos Then
            productName = Mid(htmlContent, startPos, endPos - startPos)
            productName = Trim(Replace(Replace(productName, vbCrLf, ""), vbLf, ""))
            If productName <> "" Then
                ExtractProductName = productName
                Exit Function
            End If
        End If
    End If
    
    ' Pattern 3: Generic product description search
    startPos = InStr(htmlContent, "Product:")
    If startPos > 0 Then
        startPos = startPos + 8
        endPos = InStr(startPos, htmlContent, vbCrLf)
        If endPos = 0 Then endPos = InStr(startPos, htmlContent, "<")
        If endPos > startPos Then
            productName = Mid(htmlContent, startPos, endPos - startPos)
            productName = Trim(Replace(Replace(productName, vbCrLf, ""), vbLf, ""))
            If productName <> "" Then
                ExtractProductName = productName
                Exit Function
            End If
        End If
    End If
    
    ExtractProductName = ""
End Function

' Alternative function to run for a single cell
Sub GetSingleProductName()
    Dim activeRow As Long
    activeRow = ActiveCell.Row
    
    If Cells(activeRow, "B").Value <> "" Then
        Call GetHPProductName
    Else
        MsgBox "Please select a cell in a row that has a serial number in column B"
    End If
End Sub
