using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.WrapType = WrapType.None;          // No text wrapping – the box stays in the header.
        textBox.Height = 50;                       // Height in points.
        textBox.Width = 200;                       // Width in points.
        textBox.HorizontalAlignment = HorizontalAlignment.Center;
        textBox.VerticalAlignment = VerticalAlignment.Top;

        // Add a paragraph with a run of text inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = (Paragraph)textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Header Text Box");
        para.AppendChild(run);

        // Insert the text box into the header.
        builder.InsertNode(textBox);

        // Return the cursor to the main document body.
        builder.MoveToSection(0);

        // Add enough content to generate multiple pages.
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the document. The text box will appear in the header on every page.
        doc.Save("HeaderWithTextBox.docx");
    }
}
