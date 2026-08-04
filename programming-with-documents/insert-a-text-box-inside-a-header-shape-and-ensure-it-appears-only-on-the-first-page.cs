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

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Height = 50;
        textBox.Width = 200;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph inside the text box and put some text.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "First page header text box");
        para.AppendChild(run);

        // Insert the text box into the header.
        builder.CurrentParagraph.AppendChild(textBox);

        // Return to the main document body.
        builder.MoveToSection(0);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstPageHeaderTextBox.docx");
        doc.Save(outputPath);
    }
}
