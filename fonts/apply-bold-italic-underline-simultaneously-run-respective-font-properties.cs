using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph of the document (it exists by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with the desired text.
        Run run = new Run(doc, "Bold, Italic, Underlined");

        // Apply bold, italic, and underline formatting via the Font properties.
        run.Font.Bold = true;                     // Make the text bold.
        run.Font.Italic = true;                   // Make the text italic.
        run.Font.Underline = Underline.Single;    // Apply a single underline.

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a file.
        doc.Save("FormattedRun.docx");
    }
}
