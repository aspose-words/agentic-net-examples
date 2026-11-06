using System;
using System.Drawing;
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

        // Insert an original rectangle shape.
        builder.InsertShape(ShapeType.Rectangle, 150, 80);
        // Retrieve the inserted shape.
        Shape originalShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];

        // Clone the original shape (deep clone).
        Shape clonedShape = (Shape)originalShape.Clone(true);
        // Change the fill color of the cloned shape.
        clonedShape.FillColor = Color.Blue;

        // Move the builder to a new paragraph after the original shape.
        builder.Writeln(); // creates a new paragraph.
        // Append the cloned shape to the new paragraph.
        builder.CurrentParagraph.AppendChild(clonedShape);

        // Save the document.
        string outputPath = "ClonedShape.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
