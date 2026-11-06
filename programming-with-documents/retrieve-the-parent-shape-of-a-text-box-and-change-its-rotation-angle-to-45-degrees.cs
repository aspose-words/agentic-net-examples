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

        // Insert a text box shape.
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
        // Add some text inside the text box.
        builder.Writeln("Sample text inside the text box.");

        // Retrieve the parent Shape of the text box (the shape itself in this case).
        // If we had a node inside the shape, we could use node.GetAncestor(NodeType.Shape) to get it.
        Shape parentShape = textBox; // The shape we just created is the parent.

        // Change the rotation angle to 45 degrees.
        parentShape.Rotation = 45;

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
