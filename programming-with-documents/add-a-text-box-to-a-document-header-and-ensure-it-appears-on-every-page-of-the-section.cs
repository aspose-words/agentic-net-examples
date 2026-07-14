using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsHeaderTextBoxExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a DocumentBuilder which will be used to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the primary header of the first section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert a floating text box shape into the header.
            // Width = 200 points, Height = 50 points.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
            // Ensure the shape does not wrap with surrounding text.
            textBox.WrapType = WrapType.None;

            // Add a paragraph inside the text box and center its text.
            Paragraph para = new Paragraph(doc);
            textBox.AppendChild(para);
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Add the actual text that will appear inside the text box.
            Run run = new Run(doc, "Header TextBox");
            para.AppendChild(run);

            // Return the cursor to the main body of the document.
            builder.MoveToSection(0);

            // Add some body content spanning multiple pages to demonstrate the header.
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderWithTextBox.docx");

            // Save the document to disk.
            doc.Save(outputPath);
        }
    }
}
