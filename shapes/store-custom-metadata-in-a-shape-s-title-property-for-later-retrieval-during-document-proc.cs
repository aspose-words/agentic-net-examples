using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ShapeTitleExample
{
    public static void Main()
    {
        // Define file paths in the system's temporary directory.
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeShapeDemo");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "ShapeWithTitle.docx");

        // -------------------- Create document and shape --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Store custom metadata in the Title property.
        string expectedTitle = "CustomMetadata-12345";
        shape.Title = expectedTitle;

        // Save the document.
        doc.Save(docPath);

        // Verify that the file was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("Document was not saved correctly.");

        // -------------------- Load document and retrieve title --------------------
        Document loadedDoc = new Document(docPath);
        Shape loadedShape = (Shape)loadedDoc.GetChild(NodeType.Shape, 0, true);

        // Retrieve the Title property.
        string actualTitle = loadedShape.Title;

        // Validate that the stored metadata matches the expected value.
        if (actualTitle != expectedTitle)
            throw new InvalidOperationException($"Title mismatch. Expected: '{expectedTitle}', Actual: '{actualTitle}'");

        // Optional: output confirmation (no interactive input required).
        Console.WriteLine("Shape title stored and retrieved successfully.");
    }
}
