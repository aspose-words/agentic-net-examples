using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class ApplyFontToRuns
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to work with.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run of text.
        Run run = new Run(doc, "Hello world with custom font!");

        // Apply a specific font to the run.
        run.Font.Name = "Courier New";

        // Optionally set additional font properties.
        run.Font.Size = 24;               // 24 points
        run.Font.Color = Color.Blue;      // Blue color

        // Append the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a DOCX file.
        doc.Save("AppliedFont.docx");
    }
}
