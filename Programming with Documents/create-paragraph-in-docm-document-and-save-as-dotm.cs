using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document (DOCM format by default)
        Document doc = new Document();

        // Create a new paragraph belonging to the document
        Paragraph paragraph = new Paragraph(doc);
        // Optional: set paragraph style or alignment
        paragraph.ParagraphFormat.StyleName = "Heading 1";

        // Create a run with the desired text and add it to the paragraph
        Run run = new Run(doc, "Hello from a DOTM template!");
        paragraph.AppendChild(run);

        // Append the paragraph to the body of the first section
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document as a macro‑enabled template (DOTM)
        doc.Save("Result.dotm", SaveFormat.Dotm);
    }
}
