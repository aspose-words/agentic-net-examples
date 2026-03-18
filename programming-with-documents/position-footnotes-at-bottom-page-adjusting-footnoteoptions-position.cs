using System;
using Aspose.Words;
using Aspose.Words.Notes;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text and a footnote to demonstrate the positioning.
        builder.Write("Hello world! ");
        builder.InsertFootnote(FootnoteType.Footnote, "This footnote will appear at the bottom of the page.");

        // Explicitly set footnotes to be placed at the bottom of each page.
        doc.FootnoteOptions.Position = FootnotePosition.BottomOfPage;

        // Save the resulting document.
        doc.Save("FootnoteBottomOfPage.docx");
    }
}
