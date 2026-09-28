using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the first page has a different header/footer.
        doc.FirstSection.PageSetup.DifferentFirstPageHeaderFooter = true;

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move to the first page header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);

        // Insert a text box shape into the header.
        // Width = 200 points, Height = 50 points.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 50);

        // Add text inside the text box.
        // Move the cursor inside the shape.
        builder.MoveTo(textBox.FirstParagraph);
        builder.Writeln("First page header text");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
