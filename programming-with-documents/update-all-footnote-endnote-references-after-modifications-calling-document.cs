using System;
using Aspose.Words;
using Aspose.Words.Notes;

class UpdateFootnotesAndEndnotes
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Example modification: insert a footnote at the end of the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(firstParagraph);
        builder.InsertFootnote(FootnoteType.Footnote, "Added footnote text.");

        // Update all fields in the document (e.g., REF, PAGE, etc.).
        doc.UpdateFields();

        // Update the ActualReferenceMark property for all footnotes and endnotes.
        doc.UpdateActualReferenceMarks();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
