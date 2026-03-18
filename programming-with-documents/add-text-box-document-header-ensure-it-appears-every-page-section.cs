using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a floating text box (200 pt wide, 50 pt high) into the header.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        textBox.WrapType = WrapType.None; // Prevent text wrapping around the shape.

        // Position the builder inside the text box to add its content.
        builder.MoveTo(textBox.LastParagraph);
        builder.Write("Header TextBox");

        // Return to the main body of the first section.
        builder.MoveToSection(0);

        // Add body content that spans several pages to demonstrate the header repeats.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i} body text.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        doc.Save("HeaderWithTextBox.docx");
    }
}
