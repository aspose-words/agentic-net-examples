using System;
using Aspose.Words;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Get the first paragraph of the document (it exists by default).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run of text.
        Run run = new Run(doc, "Hello world!");

        // Apply a specific font to the run.
        run.Font.Name = "Courier New";   // Set the font name.
        run.Font.Size = 24;              // Optional: set the font size.

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a DOCX file.
        doc.Save("Output.docx");
    }
}
