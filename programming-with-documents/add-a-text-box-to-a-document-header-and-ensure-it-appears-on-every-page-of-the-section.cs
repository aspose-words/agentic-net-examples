using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a text box shape into the header.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        // Optional formatting for the text box.
        textBox.WrapType = WrapType.Inline;
        textBox.FillColor = Color.LightYellow;
        textBox.StrokeColor = Color.Black;

        // Add text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Header Text Box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Return the builder to the main document body.
        builder.MoveToDocumentEnd();

        // Add enough content to generate multiple pages.
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln($"This is line {i + 1} of the document body.");
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (File.Exists(outputPath))
        {
            // Reload the document to ensure it can be opened.
            Document loadedDoc = new Document(outputPath);
        }
    }
}
