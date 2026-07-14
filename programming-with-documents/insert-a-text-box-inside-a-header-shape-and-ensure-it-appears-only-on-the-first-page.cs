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

        // Use DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable a different header for the first page.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Move the cursor to the first‑page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Ensure there is a paragraph to host the shape.
        builder.Writeln();

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;
        textBox.Height = 50;
        textBox.Width = 200;
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph inside the text box and put some text.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph tbParagraph = textBox.FirstParagraph;
        tbParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run tbRun = new Run(doc, "Header only on first page");
        tbParagraph.AppendChild(tbRun);

        // Append the text box to the current paragraph in the header.
        builder.CurrentParagraph.AppendChild(textBox);

        // Return to the main document body.
        builder.MoveToSection(0);

        // Add content spanning three pages.
        builder.Writeln("Content of the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of the second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of the third page.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderFirstTextBox.docx");
        doc.Save(outputPath);
    }
}
