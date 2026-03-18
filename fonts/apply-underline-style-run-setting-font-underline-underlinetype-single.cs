using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class UnderlineRunExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new paragraph to hold the run.
        Paragraph paragraph = new Paragraph(doc);

        // Create a run with the desired text.
        Run run = new Run(doc, "This text will be underlined.");

        // Apply a single underline to the run's font.
        run.Font.Underline = Underline.Single;

        // Add the run to the paragraph, and the paragraph to the document body.
        paragraph.AppendChild(run);
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document to disk.
        doc.Save("UnderlineRun.docx");
    }
}
