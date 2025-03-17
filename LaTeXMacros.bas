Attribute VB_Name = "LaTeXMacros"
' Attribute VB_Name = "LaTeXMacros"
'Copyright (C) 2007 Tyler A. Davis
'Copyright (C) 2007 Philip Stevenson
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

Sub LaTeXEntry()
'
' LaTeXEntry Macro
' Creates an equation image using LaTeX
'
    ' Check selection for useful data
    If (Selection.Type = wdSelectionInlineShape) Then ' Image?
        tex_src = Selection.InlineShapes(1).AlternativeText
    ElseIf (Selection.Type <> wdSelectionIP) Then ' Not an Insertion Point?
        tex_src = Selection.Text
    End If

    ' Initialize the font size combo box
    LaTeX_Entry.ComboFontSize.AddItem ("10")
    LaTeX_Entry.ComboFontSize.AddItem ("11")
    LaTeX_Entry.ComboFontSize.AddItem ("12")

'   If ((Selection.Font.Size >= 10.5) And (Selection.Font.Size < 11.5)) Then
'     LaTeX_Entry.ComboFontSize.Value = "11"
'   ElseIf (Selection.Font.Size >= 11.5) Then
'     LaTeX_Entry.ComboFontSize.Value = "12"
'   Else
      '10 point is the default
      LaTeX_Entry.ComboFontSize.Value = "10"
'   End If

    If tex_src = "" Then
        
        sample = "\documentclass{article}" & vbCrLf
        sample = sample & "\usepackage{amsmath}" & vbCrLf
        sample = sample & "\pagestyle{empty}" & vbCrLf
        sample = sample & "\begin{document}" & vbCrLf & vbCrLf
        sample = sample & "\[" & vbCrLf
        sample = sample & "y=ax+b" & vbCrLf
        sample = sample & "\]" & vbCrLf & vbCrLf
        sample = sample & "\end{document}"
    
        LaTeX_Entry.Entry_Box.Text = sample
    
    ElseIf InStr(tex_src, "documentclass") > 0 Then

        LaTeX_Entry.Entry_Box.Text = tex_src

    End If

    ' Start the official dialog
    Load LaTeX_Entry
    LaTeX_Entry.Show
End Sub

Sub EqnNumber()
'
' EqnNumber Macro
' Adds a bookmarked equation number.
'
    ' Get the bookmark data from the user
    Load EqnBookmark
    EqnBookmark.Show
End Sub

Sub ResetTextPosition()
'
' ResetTextPosition Macro
' Sets the vertical position of the selected text to 0
    
    Selection.Font.Position = 0
End Sub


