using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank DOCM document.
        Document doc = new Document();

        // Create a new paragraph belonging to the document.
        Paragraph paragraph = new Paragraph(doc);

        // Add a run (text) to the paragraph.
        Run run = new Run(doc, "This is a new paragraph in a DOCM document.");
        paragraph.AppendChild(run);

        // Append the paragraph to the body of the first section.
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document as a macro‑enabled template (DOTM).
        doc.Save("Result.dotm", SaveFormat.Dotm);
    }
}
