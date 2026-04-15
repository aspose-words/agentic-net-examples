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

        // Move the builder cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a floating text box shape into the header.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        // Prevent the text box from affecting the surrounding text.
        textBox.WrapType = WrapType.None;
        // Optional: set a border for visual clarity.
        textBox.StrokeColor = System.Drawing.Color.Black;
        textBox.StrokeWeight = 0.5;

        // Add a paragraph inside the text box and write some text.
        Paragraph para = textBox.FirstParagraph;
        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Run run = new Run(doc, "Header Text Box");
        para.AppendChild(run);

        // Return to the main document body.
        builder.MoveToSection(0);

        // Add enough content to generate multiple pages.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Page {i} content.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderTextBox.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
