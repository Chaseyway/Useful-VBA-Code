Attribute VB_Name = "JSE_Shortcuts"
Sub JSE_Shortcuts()
  'Note that file path of chrome will vary per user, this must be updated when using script
  'Also note that script depends on current structure of JSE website, if this is changed the URLs in the script must be updated
Dim infotype, ticker As String

ticker = InputBox("Enter stock ticker:")
infotype = InputBox("What data do you want?" & vbCrLf & "Audited: a" & vbCrLf & "Annual Report: an" & vbCrLf & "Quarter: q" & vbCrLf & "Profile: p" & vbCrLf & "Current Combined Qoute: cq" & vbCrLf & "Price History: ph" & vbCrLf & "Stock News: n" & vbCrLf & "Stock Market News: news")

If infotype = "a" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/?tag=" & ticker & "&category_name=audited-financial-statements")

ElseIf infotype = "an" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/?tag=" & ticker & "&category_name=annual-reports")

ElseIf infotype = "q" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/?tag=" & ticker & "&category_name=quarterly-financial-statements")

ElseIf infotype = "p" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/trading/instruments/?instrument=" & ticker & "-jmd")

ElseIf infotype = "cq" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/trading/trade-quotes/?market=combined-market")

ElseIf infotype = "n" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/?tag=" & ticker)

ElseIf infotype = "news" Then
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/news/")

ElseIf infotype = "ph" Then
Dim startdate, enddate As String
startdate = InputBox("Input Start date (YYYY-MM-DD)")
enddate = InputBox("Input End date (YYYY-MM-DD)")
chromeFileLocation = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
Shell (chromeFileLocation & "-url " & "https://www.jamstockex.com/trading/instruments/price-history/?instrument=" & ticker & "-jmd&fromDate=" & startdate & "&thruDate=" & enddate)

End If

End Sub

