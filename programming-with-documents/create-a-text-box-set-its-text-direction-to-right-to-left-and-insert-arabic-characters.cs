using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text box shape.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 200, 100);

        // Create a paragraph that will hold the Arabic text.
        Paragraph para = new Paragraph(doc);
        // Set the paragraph direction to right‑to‑left.
        para.ParagraphFormat.Bidi = true;

        // Add the Arabic text run.
        Run run = new Run(doc, "مرحبا بالعالم");
        para.AppendChild(run);

        // Insert the paragraph into the text box.
        textBoxShape.AppendChild(para);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
