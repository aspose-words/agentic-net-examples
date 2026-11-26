using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the first page has a different header/footer.
        Section section = doc.FirstSection;
        section.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Create a header that will appear only on the first page.
        HeaderFooter firstPageHeader = new HeaderFooter(doc, HeaderFooterType.HeaderFirst);
        section.HeadersFooters.Add(firstPageHeader);

        // Use a DocumentBuilder to work inside the first‑page header.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Insert a text box shape into the header.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);
        textBox.WrapType = WrapType.None;
        textBox.StrokeColor = Color.Black;
        textBox.FillColor = Color.LightYellow;

        // Add text inside the text box using the builder.
        builder.MoveTo(textBox);
        builder.Writeln("First page header text box");

        // Save the document.
        string outputPath = "HeaderTextBox.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
        }
    }
}
