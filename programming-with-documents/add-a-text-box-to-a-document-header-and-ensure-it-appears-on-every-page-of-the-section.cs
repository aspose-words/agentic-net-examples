using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the primary header exists.
        Section section = doc.FirstSection;
        HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header == null)
        {
            header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            section.HeadersFooters.Add(header);
        }

        // Move the builder to the header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a textbox shape into the header.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        textBox.WrapType = WrapType.Inline;

        // Add text to the textbox.
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, "Header TextBox Content");
        para.AppendChild(run);
        textBox.AppendChild(para);

        // Add body content to generate multiple pages.
        builder.MoveToDocumentEnd();
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the document.
        string outputFile = "HeaderWithTextBox.docx";
        doc.Save(outputFile);
        Console.WriteLine($"Document saved as {outputFile}");
    }
}
