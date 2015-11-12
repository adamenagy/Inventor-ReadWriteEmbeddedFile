# ReadWriteEmbeddedFile
Shows how you can read/write files embedded inside Inventor

If you have embedded files inside Inventor and they are not e.g. an Excel file where you could use the Excel API to interact with the file, then it's a bit difficult to access it.  

E.g. if you have a text file embedded in an Inventor document, then by default when you try to programmatically access that file through ReferencedOLEFileDescriptor.Activate(kShowOLEVerb, Nothing), then it will just bring up Notepad with the file opened in it.  

It seems that when Inventor opens a document with embedded files in it then it automatically creates files for them on the disk in the temp folder: C:\Users\<user name>\AppData\Local\Temp  
When you try to edit the files then the path of the temp file will be passed to the editor and then when the editor is closed the changes get stored in the Inventor document.  

We'll hook into this mechanism using the executable in the repo to edit embedded text files. The console app has 3 command line options:
* /r: use it to register the app in the registry for text file opening, so that it will be used when Inventor is opening the embedded text file for edit  
* /u: use it to unregister the app. It will change the registry values back to what they were  
* &lt;file path&gt;: when providing the path to the file it will add a new line to it with the current date and time  

You can use the compiled application with an Inventor part document that has an embedded text document in it. When using the below code the text file you want to modify needs to be selected in the Inventor model browser and the path of the exe needs to be set:

```VB.NET
Sub ModifyEmbeddedTextFile()
  Dim doc As PartDocument
  Set doc = ThisApplication.ActiveDocument
  
  Dim obj As ReferencedOLEFileDescriptor
  Set obj = doc.SelectSet(1)
  
  ' Register our reader for text files
  Call Shell( _
    """" & _
    "C:\Temp\TextReaderCS.exe" & _
    """ """ & _
    "/r" & _
    """", vbNormalFocus)
  
  ' Get the embedded object opened
  ' Note: this will create a temporary file with
  ' the content of the embedded object
  ' the path of which will be passed to the editor
  ' program
  ' This call will return once the editor program
  ' is closed
  Call obj.Activate(kShowOLEVerb, Nothing)
  
  ' Unregister our editor for text files
  Call Shell( _
    """" & _
    "C:\Temp\TextReaderCS.exe" & _
    """ """ & _
    "/u" & _
    """", vbNormalFocus)
End Sub
```
