using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ExtractOleImages
{
    public static void Main()
    {
        // Define file and folder names.
        const string docPath = "DocumentWithOle.docx";
        const string oleDataFile = "sample.txt";

        // Create a simple text file to embed as an OLE object.
        File.WriteAllText(oleDataFile, "This is sample OLE embedded content.");

        // Create a new Word document and embed the OLE object as an icon.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Read the sample file into a memory stream.
        using (MemoryStream oleStream = new MemoryStream(File.ReadAllBytes(oleDataFile)))
        {
            // Insert the OLE object. The 'asIcon' flag is true so the shape will contain an image.
            builder.InsertOleObject(oleStream, "Package", true, null);
        }

        // Save the document that now contains an OLE object.
        doc.Save(docPath);

        // Load the document back (demonstrates the load step).
        Document loadedDoc = new Document(docPath);

        // Get all shape nodes in the document.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Process only OLE object shapes that have an image (icon).
            if (shape.ShapeType == ShapeType.OleObject && shape.HasImage)
            {
                OleFormat oleFormat = shape.OleFormat;
                // Use the ProgId as part of the file name; replace characters that are invalid in file names.
                string progId = oleFormat.ProgId ?? "OleObject";
                foreach (char c in Path.GetInvalidFileNameChars())
                    progId = progId.Replace(c, '_');

                // Determine the appropriate image file extension.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{progId}_icon{extension}";
                string imagePath = Path.Combine(Directory.GetCurrentDirectory(), imageFileName);

                // Save the image (icon) to the file system.
                shape.ImageData.Save(imagePath);
                extractedCount++;

                // Optional: verify that the file was created.
                if (!File.Exists(imagePath))
                    throw new InvalidOperationException($"Failed to save extracted image to '{imagePath}'.");
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No OLE object images were found and extracted.");

        // Clean up temporary files used for the example.
        File.Delete(oleDataFile);
    }
}
