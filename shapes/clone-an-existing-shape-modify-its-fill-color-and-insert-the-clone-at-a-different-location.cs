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

        // Insert an original rectangle shape inline and set its fill color.
        Shape originalShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        originalShape.FillColor = Color.LightBlue;

        // Clone the original shape (deep clone) and change the fill color of the clone.
        Shape clonedShape = (Shape)originalShape.Clone(true);
        clonedShape.FillColor = Color.LightCoral;

        // Insert the cloned shape at a different location.
        // Create a new paragraph and append the cloned shape to it.
        Paragraph newParagraph = new Paragraph(doc);
        doc.FirstSection.Body.AppendChild(newParagraph);
        newParagraph.AppendChild(clonedShape);

        // Save the document to disk.
        string outputPath = "ClonedShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optional: inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
