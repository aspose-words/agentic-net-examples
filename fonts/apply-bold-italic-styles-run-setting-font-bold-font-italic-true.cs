using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Get the first paragraph of the document (it always exists in a new document).
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a run with the desired text.
        Run run = new Run(doc, "Bold and Italic text");

        // Apply bold and italic formatting to the run.
        run.Font.Bold = true;
        run.Font.Italic = true;

        // Add the run to the paragraph.
        paragraph.AppendChild(run);

        // Save the document to a file.
        doc.Save("BoldItalicRun.docx");
    }
}
