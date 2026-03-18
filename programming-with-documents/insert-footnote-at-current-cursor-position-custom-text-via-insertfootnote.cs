using System;
using Aspose.Words;
using Aspose.Words.Notes;

class InsertFootnoteExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text where the footnote will be attached.
        builder.Write("This sentence will have a footnote.");

        // Insert a footnote at the current cursor position with custom text.
        // FootnoteType.Footnote specifies that we are inserting a footnote (not an endnote).
        builder.InsertFootnote(FootnoteType.Footnote, "This is the custom footnote text.");

        // Save the document to a file.
        doc.Save("FootnoteExample.docx");
    }
}
