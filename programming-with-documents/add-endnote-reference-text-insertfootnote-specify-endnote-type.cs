using System;
using Aspose.Words;
using Aspose.Words.Notes;

class EndnoteExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text that will be referenced by the endnote.
        builder.Write("This text will be referenced by an endnote.");

        // Insert an endnote with the desired reference text.
        // FootnoteType.Endnote tells the builder to create an endnote rather than a footnote.
        builder.InsertFootnote(FootnoteType.Endnote, "Endnote comment regarding referenced text.");

        // Save the document to a file.
        doc.Save("EndnoteExample.docx");
    }
}
