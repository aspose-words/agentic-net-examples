using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX containing vector shapes.
        string docPath = "sample.docx";
        CreateSampleDocument(docPath);

        // Load the document that contains the vector shapes.
        Document doc = new Document(docPath);

        // Prepare the output folder for the extracted EMF files.
        string outputDir = "ExtractedEMF";
        Directory.CreateDirectory(outputDir);

        // Find all Shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        // Iterate through each shape.
        for (int i = 0; i < shapeNodes.Count; i++)
        {
            Shape shape = (Shape)shapeNodes[i];

            // We are interested only in vector shapes (they do not contain an image).
            if (!shape.HasImage)
            {
                // Create a temporary document that will hold only the cloned shape.
                Document tempDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(tempDoc);

                // Import the shape into the temporary document before inserting.
                Node importedShape = tempDoc.ImportNode(shape, true);
                builder.InsertNode(importedShape);

                // Save the temporary document as an EMF image.
                string outPath = Path.Combine(outputDir, $"vector-{extractedCount}.emf");
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf);
                tempDoc.Save(outPath, saveOptions);

                // Validate that the EMF file was created.
                if (!File.Exists(outPath))
                    throw new Exception($"Failed to save EMF file: {outPath}");

                extractedCount++;
            }
        }

        // Ensure that at least one vector shape was extracted.
        if (extractedCount == 0)
            throw new Exception("No vector images were extracted.");

        Console.WriteLine($"Extracted {extractedCount} vector shape(s) to EMF format in folder '{outputDir}'.");
    }

    // Helper method to create a sample document with vector shapes.
    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a star shape (vector).
        builder.InsertShape(ShapeType.Star, 100, 100);
        builder.Writeln();

        // Insert an arrow shape (vector).
        builder.InsertShape(ShapeType.Arrow, 150, 50);
        builder.Writeln();

        // Save the document to the specified path.
        doc.Save(path);
    }
}
