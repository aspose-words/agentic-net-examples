using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document (lifecycle rule)
        Document doc = new Document();

        // Create a new paragraph belonging to the document
        Paragraph paragraph = new Paragraph(doc);
        // Example formatting: center alignment
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Create a run with text and add it to the paragraph
        Run run = new Run(doc, "Hello from DOCM!");
        // Example formatting: blue font color
        run.Font.Color = Color.Blue;
        paragraph.AppendChild(run);

        // Append the paragraph to the body of the first section
        doc.FirstSection.Body.AppendChild(paragraph);

        // Save the document as a macro-enabled DOCM file (save rule)
        doc.Save("CreatedDocument.docm", SaveFormat.Docm);
    }
}
