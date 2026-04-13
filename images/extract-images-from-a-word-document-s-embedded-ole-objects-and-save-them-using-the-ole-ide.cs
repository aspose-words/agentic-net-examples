using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class ExtractOleImages
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample file that will be embedded as an OLE object.
        string sampleFilePath = Path.Combine(outputDir, "sample.txt");
        File.WriteAllText(sampleFilePath, "This is a sample text file embedded as OLE.");

        // Create a new document and embed the OLE object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        using (FileStream oleStream = File.OpenRead(sampleFilePath))
        {
            // Insert the OLE object as a package (generic container).
            // The icon will be generated automatically.
            builder.InsertOleObject(oleStream, "Package", false, null);
        }

        // Save the document (optional, just for reference).
        string docPath = Path.Combine(outputDir, "DocumentWithOle.docx");
        doc.Save(docPath);

        // Load the document and extract images (icons) from OLE objects.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int index = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that contain an OLE object.
            if (shape.OleFormat == null)
                continue;

            // Determine a safe identifier for the OLE object.
            string progId = !string.IsNullOrEmpty(shape.OleFormat.ProgId) ? shape.OleFormat.ProgId : "Package";

            // Save the OLE data itself (e.g., the embedded file).
            string oleFileName = $"{progId}_{index}{shape.OleFormat.SuggestedExtension}";
            string oleFullPath = Path.Combine(outputDir, oleFileName);
            shape.OleFormat.Save(oleFullPath);

            // If the OLE shape has an associated image (icon), extract it.
            if (shape.HasImage)
            {
                string imageExtension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{progId}_{index}{imageExtension}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                extractedCount++;
            }

            index++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No OLE icon images were extracted from the document.");

        // Optionally, inform about the result (no interactive prompts required).
        Console.WriteLine($"Extraction complete. {extractedCount} image(s) saved to '{outputDir}'.");
    }
}
