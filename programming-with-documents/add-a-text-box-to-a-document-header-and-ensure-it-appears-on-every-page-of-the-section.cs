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

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Create a floating text box shape.
        Shape textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = 200;               // Width in points.
        textBox.Height = 50;               // Height in points.
        textBox.WrapType = WrapType.None;  // Prevent text wrapping.
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        textBox.Left = 50;                 // Position from the left edge of the page.
        textBox.Top = 20;                  // Position from the top edge of the page.

        // Add a paragraph with a run of text inside the text box.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Header TextBox");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert the text box into the header.
        builder.InsertNode(textBox);

        // Return to the main body of the document.
        builder.MoveToSection(0);
        builder.Writeln("First page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page content.");

        // Save the document to a file in the current directory.
        doc.Save("HeaderWithTextBox.docx");
    }
}
