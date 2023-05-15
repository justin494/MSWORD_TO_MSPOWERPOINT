import win32com.client as w

def copy_content(word, ppt, word_doc_path, ppt_pres_path):
# Open MS Word and MS PowerPoint applications
    word_app = w.Dispatch('Word.Application')
    ppt_app = w.Dispatch('Powerpoint.Application')
    # Open MS Word document and select content to be copied
    word_doc = word_app.Documents.Open(word_doc_path)
    word_range = word_doc.Range(0, 0)
    word_range.WholeStory()
    word_range.Copy()

    # Open MS PowerPoint presentation
    ppt_pres = ppt_app.Presentations.Open(ppt_pres_path)

# Add a new slide if content does not fit in one slide
    if len(ppt_pres.Slides) > 0:
        last_slide = ppt_pres.Slides[len(ppt_pres.Slides)]
        last_shape = last_slide.Shapes[len(last_slide.Shapes)]
        if last_shape.HasTextFrame:
            if last_shape.TextFrame.TextRange.Text != '':
                ppt_slide = ppt_pres.Slides.Add(ppt_pres.Slides.Count + 1, 12)

    # Paste the content into the new slide
    ppt_slide = ppt_pres.Slides.Add(ppt_pres.Slides.Count + 1, 12)
    ppt_shape = ppt_slide.Shapes.PasteSpecial(DataType=2)
    ppt_shape.Left = 25
    ppt_shape.Top = 25

    # Save and close MS PowerPoint presentation and MS Word document
    ppt_pres.SaveAs('D:\Assignment digiplus\Sample_out_pptx.pptx')
    ppt_pres.Close()
    word_doc.Close()

copy_content('Word.Application', 'Powerpoint.Application', 'D:\Assignment digiplus\Sample_DOCX.docx', 'D:\Assignment digiplus\Sample_PPTX.pptx')
    