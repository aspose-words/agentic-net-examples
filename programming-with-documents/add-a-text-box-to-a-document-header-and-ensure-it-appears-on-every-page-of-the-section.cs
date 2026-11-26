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

        // Add enough content to generate several pages.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln($"This is page {i + 1} content.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Move the builder to the primary header of the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Create a text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;
        textBox.Height = 50;
        textBox.WrapType = WrapType.Inline;

        // Optional: make the text box fit its contents.
        textBox.TextBox.FitShapeToText = true;

        // Add a paragraph with text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Header Text Box");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the header.
        builder.InsertNode(textBox);

        // Save the document.
        string outputPath = "HeaderWithTextBox.docx";
        doc.Save(outputPath);
    }
}
