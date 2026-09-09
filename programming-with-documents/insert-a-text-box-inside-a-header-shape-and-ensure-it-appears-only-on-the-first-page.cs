using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Height = 50;
        textBox.Width = 200;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph and a run of text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "First page header text box");
        para.AppendChild(run);

        // Insert the text box into the header.
        builder.InsertNode(textBox);

        // Add content to generate multiple pages.
        builder.MoveToSection(0);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FirstPageHeaderTextBox.docx");
        doc.Save(outputPath);
    }
}
