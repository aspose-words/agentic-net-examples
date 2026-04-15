using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the new image to embed and the output document.
        string newImagePath = Path.Combine("Data", "NewImage.png");
        string outputDocPath = Path.Combine("Output", "ReplacedOle.docx");

        // Ensure the output directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputDocPath));

        // -----------------------------------------------------------------
        // 1. Create a new document and insert a placeholder OLE object.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a dummy OLE package (empty memory stream) so that we have an OLE object to replace.
        using (MemoryStream dummyStream = new MemoryStream())
        {
            // Insert as a regular (non‑icon) OLE object using the generic "Package" ProgID.
            Shape placeholderOle = builder.InsertOleObject(dummyStream, "Package", false, null);
            // Optionally set a display name for the placeholder.
            placeholderOle.OleFormat.OlePackage.FileName = "Placeholder.bin";
            placeholderOle.OleFormat.OlePackage.DisplayName = "Placeholder";
        }

        // -----------------------------------------------------------------
        // 2. Locate the first OLE object shape in the document.
        // -----------------------------------------------------------------
        Shape oldOleShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (oldOleShape == null || oldOleShape.ShapeType != ShapeType.OleObject)
        {
            Console.WriteLine("No OLE object found to replace.");
            return;
        }

        // Move the builder cursor to the existing OLE shape.
        builder.MoveTo(oldOleShape);

        // -----------------------------------------------------------------
        // 3. Insert the new image as an OLE object.
        // -----------------------------------------------------------------
        if (!File.Exists(newImagePath))
        {
            Console.WriteLine($"Image file not found: {newImagePath}");
            return;
        }

        using (FileStream imageStream = File.OpenRead(newImagePath))
        {
            // Insert the image as an OLE object (displayed as its content, not as an icon).
            builder.InsertOleObject(imageStream, "Package", false, null);
        }

        // -----------------------------------------------------------------
        // 4. Remove the original placeholder OLE object.
        // -----------------------------------------------------------------
        oldOleShape.Remove();

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputDocPath);
        Console.WriteLine($"Document saved to: {outputDocPath}");
    }
}
