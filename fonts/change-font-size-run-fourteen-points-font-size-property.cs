using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph to host the run.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some sample text.
        Run run = new Run(doc, "Sample text with 14‑point font.");

        // Set the font size of the run to fourteen points.
        run.Font.Size = 14;

        // Append the run to the paragraph.
        para.AppendChild(run);

        // Define an output folder (adjust as needed).
        string artifactsDir = @"C:\Temp\";

        // Save the document.
        doc.Save(artifactsDir + "RunFontSize.docx");
    }
}
