using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class CreateDocmParagraph
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one section, body and paragraph.
        // (A newly created Document already contains these nodes, but calling EnsureMinimum
        // guarantees the structure if the document was altered earlier.)
        doc.EnsureMinimum();

        // Access the body of the first (and only) section.
        Body body = doc.FirstSection.Body;

        // Create a new paragraph belonging to the document.
        Paragraph paragraph = new Paragraph(doc);
        // Set paragraph formatting (optional).
        paragraph.ParagraphFormat.StyleName = "Heading 1";
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Create a run with the desired text.
        Run run = new Run(doc);
        run.Text = "Hello World!";
        run.Font.Color = Color.Red;

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Append the paragraph to the document body.
        body.AppendChild(paragraph);

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("Output.docm", SaveFormat.Docm);
    }
}
