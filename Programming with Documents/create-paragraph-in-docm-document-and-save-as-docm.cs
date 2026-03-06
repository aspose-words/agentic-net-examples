using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one section and body.
        doc.EnsureMinimum();

        // Create a new paragraph that belongs to the document.
        Paragraph para = new Paragraph(doc);
        para.ParagraphFormat.StyleName = "Heading 1";
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Append the paragraph to the body of the first section.
        doc.FirstSection.Body.AppendChild(para);

        // Create a run with text, set its color, and add it to the paragraph.
        Run run = new Run(doc, "Hello World!");
        run.Font.Color = Color.Red;
        para.AppendChild(run);

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("Result.docm", SaveFormat.Docm);
    }
}
