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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Praesent commodo cursus magna, vel scelerisque nisl consectetur et.");

        // Insert a floating rectangle shape.
        // The shape is positioned 100 points from the left and top of the page,
        // has a size of 100x100 points, and uses Square wrapping (text wraps on both sides).
        Shape shape = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            100, 100,
            WrapType.Square);

        // Make the shape visible with a light gray fill.
        shape.FillColor = Color.LightGray;

        // Add more text that will wrap around the shape.
        builder.Writeln("\n" + new string('A', 500));

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WrapShape.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
