using System;
using Aspose.Words;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Get the first paragraph (created by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with some text.
        Run run = new Run(doc, "Hello world!");

        // Assign a blue color to the run's font.
        run.Font.Color = Color.Blue;

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document.
        doc.Save("RunBlueColor.docx");
    }
}
