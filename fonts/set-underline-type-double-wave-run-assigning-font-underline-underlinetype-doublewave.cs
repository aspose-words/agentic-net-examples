using System;
using Aspose.Words;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Get the first paragraph (created by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with sample text.
        Run run = new Run(doc, "This text has a double‑wave underline.");

        // Set the underline type to double wave.
        run.Font.Underline = Underline.WavyDouble; // double wave underline

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document.
        doc.Save("RunDoubleWaveUnderline.docx");
    }
}
