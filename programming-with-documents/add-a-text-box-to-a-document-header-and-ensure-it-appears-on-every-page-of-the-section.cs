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

        // Move the cursor to the primary header of the first (and only) section.
        // The primary header appears on every page unless different headers are enabled.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a floating text box shape into the header.
        // Width = 200 points, Height = 50 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        // Ensure the text box does not wrap with surrounding text.
        textBox.WrapType = WrapType.None;

        // Add a paragraph inside the text box.
        textBox.AppendChild(new Paragraph(doc));
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Add the desired text to the text box.
        Run run = new Run(doc, "Header Text Box");
        para.AppendChild(run);

        // Return the builder to the main document body.
        builder.MoveToSection(0);

        // Add some content to generate multiple pages.
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the document to a file in the same folder as the executable.
        doc.Save("HeaderWithTextBox.docx");
    }
}
