using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeTypeEnumerator
{
    public static void Main()
    {
        // Define output folder and file names.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "SampleShapes.docx");

        // Create a new document and insert a few shapes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Insert a floating ellipse.
        builder.InsertShape(ShapeType.Ellipse, RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 150, 80, 80, WrapType.None);

        // Insert a text box.
        builder.InsertShape(ShapeType.TextBox, 120, 60);

        // Save the document.
        doc.Save(docPath);

        // Verify that the document was saved.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("The document was not saved correctly.", docPath);

        // Load the saved document.
        Document loadedDoc = new Document(docPath);

        // Retrieve all shape nodes in the document.
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .OfType<Shape>();

        // Output each shape's type to the console.
        foreach (Shape shape in shapes)
        {
            Console.WriteLine(shape.ShapeType);
        }
    }
}
