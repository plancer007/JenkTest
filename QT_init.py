import os
import sys
import win32com.client

pdf = win32com.client.Dispatch("Word.Application")  # Create new Word Object
pdf.Visible = 0  # Word Application should`t be visible
worddoc = pdf.Documents.Add()  # Create new Document Object
worddoc.PageSetup.Orientation = 1  # Make some Setup to the Document:
worddoc.PageSetup.LeftMargin = 20
worddoc.PageSetup.TopMargin = 20
worddoc.PageSetup.BottomMargin = 20
worddoc.PageSetup.RightMargin = 20
worddoc.Content.Font.Size = 11
worddoc.Content.Paragraphs.TabStops.Add(100)
worddoc.Content.Text = "Hello World!"
worddoc.Content.MoveEnd
worddoc.Close()  # Close the Word Document (a save-Dialog pops up)
pdf.Quit()  # Close the Word Application
