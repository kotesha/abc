'creating own sync point

oTimeout=100
For oTime=1 to oTimeout
          oPropval=window("Flight Reservation").WinButton("Delete Order").GetROProperty("enabled")
          If oPropval=true Then
                   Exit for
          else
                   wait(1)
          End If
Next

'Shared Object Repository associating through script

RepPath=".tsr"

RepositoriesCollection.Add(RepPath)
'Adding Repository through script at runtime

RepositoriesCollection.RemoveAll()
'For removing all repositories at run time


'regular expression for Email
Dim emailAddress
emailAddress = InputBox("emailAddress")
If emailAddress <> "" then   
'checks whether email address is empty or not
  emailAddress = Cstr(emailAddress) 
  ' if not empty takes the email address from inputbox or 
  
    blnValidEmail = RegExpTest(emailAddress)
    If blnValidEmail then
      Response.Write("Valid email address")
    Else
      Response.Write("Not a valid email address")
    End If
	Function RegExpTest(sEmail)
  RegExpTest = false
  Dim regEx, retVal
  Set regEx = New RegExp

  ' Create regular expression:
  regEx.Pattern ="^[0-9a-z*()_+%$#@!^&*=+./\}{]+@[a-z]+\.[a-z]{2,3}$"	

  ' Set pattern:
  regEx.IgnoreCase = true

  ' Set case sensitivity.
  retVal = regEx.Test(sEmail)

  ' Execute the search test.
  If not retVal Then
    exit function
  End If

  RegExpTest = true
End Function
Else
	MsgBox "Email Address Cannot be Empty"
End if

'to check checkbox is checked if not check the checkbox
var=Browser("Your Store").Page("Register Account").WebCheckBox("agree").GetROProperty("Checked")
If var=="ON" then
	Browser("Your Store").Page("Register Account").WebButton("Continue").Click
Else
	Browser("Your Store").Page("Register Account").WebCheckBox("agree").Set "ON"
End If

