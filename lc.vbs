'
' lc.vbs
'

' Script that modifies a text file.
' Be aware that this script will remove the original file
' and create a new one that will contain modifications.
' Created on 2014-11-01.
' Copyright (C) 2014 Eriks Ocakovskis.

' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

' Usage:
' All command line arguments are required and must be surrounded by double quotes
' Argument 1 - option either -r to replace text or -a to append text.
' Argument 2 - full path to a file you want to modify.
' Argument 3 - regular expression pattern of text you wan to replace
'              or after witch you want to append new text.
' Argument 4 - replacement or addition text.

Set objFS = CreateObject("Scripting.FileSystemObject")
' Command line argument 1 - option
argOption = WScript.Arguments.Item(0)

If argOption <> "-r" AND argOption <> "-a" Then
  WScript.Echo "Option you selected does not exist, use '-r' to replace or '-a' to append"
  WScript.quit
End If

' Command line argument 2 - file
argFile = WScript.Arguments.Item(1)
strTmpFile = argFile + ".temp"

If objFS.FileExists(strTmpFile) Then
  objFS.DeleteFile strTmpFile
End If

objFS.MoveFile argFile, strTmpFile

Set objRegExp = New RegExp
objRegExp.Global = True
objRegExp.IgnoreCase = True
' Command line argument 3 - pattern
objRegExp.Pattern = WScript.Arguments.Item(2)

Set objFile = objFS.OpenTextFile(strTmpFile)
Set objNewFile = objFS.OpenTextFile(argFile,2,True)
Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    If objRegExp.Test(strLine) = True Then
      If argOption= "-a" Then
        ' Command line argument 4 - addition
        strLine = objRegExp.Replace(strLine, strLine + vbCrLf + WScript.Arguments.Item(3))
      ElseIf argOption= "-r" Then
        ' Command line argument 4 - replacement
        strLine = objRegExp.Replace(strLine, WScript.Arguments.Item(3))
      End If
    ' WScript.Echo strLine
    End If
    objNewFile.WriteLine strLine
Loop

objFile.close
objNewFile.close

objFS.DeleteFile strTmpFile
