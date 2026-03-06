using System;
using System.Drawing;
using Aspose.Words;

class InsertParagraphExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph at the current cursor position.
        Paragraph newParagraph = builder.InsertParagraph();

        // Create a run with the desired text.
        Run run = new Run(doc, "Hello Aspose.Words!");

        // Apply formatting to the run (optional).
        run.Font.Color = Color.Blue;

        // Add the run to the newly inserted paragraph.
        newParagraph.AppendChild(run);

        // Save the document as a DOCX file.
        doc.Save("InsertedParagraph.docx");
    }
}
