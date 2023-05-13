import win32com.client as w

# We will first open the MS Word and MS PowerPoint application in our system without any user intervention using win32com.client library
word = w.Dispatch('Word.Application')
ppt = w.Dispatch('Powerpoint.Application')

# Now we will open the MS Word document named "DIGIPPLUS_MSWORD" and select the content to be copied
word_doc = word.Documents.Open('D:\Assignment digiplus\DIGIPPLUS_MSWORD.docx')
word_range = word_doc.Range(0, 0)
word_range.WholeStory()
word_range.Copy()

# Now we will open the MS PowerPoint named "DIGIPPLUS_MSPOWERPOINT" presentation and paste the content
ppt_pres = ppt.Presentations.Open('D:\Assignment digiplus\ok.pptx')
ppt_slide = ppt_pres.Slides.Add(1, 12) # this will add a new slide in the presentation
ppt_shape = ppt_slide.Shapes.PasteSpecial(DataType=2) # pastes the content of the clipboard into the new slide
ppt_shape.Left = 25
ppt_shape.Top = 25

# Now we will save and close the MS PowerPoint presentation and MS Word document
ppt_pres.SaveAs('D:\Assignment digiplus\ok1.pptx')
ppt_pres.Close()
word_doc.Close()

