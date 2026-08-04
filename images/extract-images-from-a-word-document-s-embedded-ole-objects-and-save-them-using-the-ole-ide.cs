using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExtractOleImages
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample icon image that will be used as the OLE object's presentation.
        string iconPath = Path.Combine(outputDir, "icon.png");
        CreateSampleImage(iconPath, 100, 100);

        // Create a sample file that will be embedded as an OLE object.
        string oleDataPath = Path.Combine(outputDir, "sample.txt");
        File.WriteAllText(oleDataPath, "This is sample OLE content.");

        // Build a document and embed the OLE object with the custom icon.
        string docPath = Path.Combine(outputDir, "DocumentWithOle.docx");
        CreateDocumentWithOle(docPath, oleDataPath, iconPath);

        // Load the document and extract images from embedded OLE objects.
        ExtractImagesFromOleObjects(docPath, outputDir);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a deterministic bitmap.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.LightGray);
        // Optionally draw something simple.
        graphics.Dispose();
        bitmap.Save(filePath);
        bitmap.Dispose();
    }

    private static void CreateDocumentWithOle(string docPath, string oleFilePath, string iconFilePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before the OLE object.
        builder.Writeln("Below is an embedded OLE object with a custom icon:");

        // Insert the OLE object using streams.
        using (FileStream oleStream = File.OpenRead(oleFilePath))
        using (FileStream iconStream = File.OpenRead(iconFilePath))
        {
            // progId "Package" works for generic files.
            // asIcon = true to display the provided icon.
            builder.InsertOleObject(oleStream, "Package", true, iconStream);
        }

        // Save the document.
        doc.Save(docPath);
    }

    private static void ExtractImagesFromOleObjects(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);

        // Get all shape nodes.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            // Identify OLE objects that have an image (icon).
            if (shape.OleFormat != null && shape.HasImage)
            {
                OleFormat ole = shape.OleFormat;
                string progId = ole.ProgId ?? "OleObject";

                // Sanitize progId for file name.
                foreach (char c in Path.GetInvalidFileNameChars())
                    progId = progId.Replace(c, '_');

                // Determine image extension based on its type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"{progId}{extension}";
                string imagePath = Path.Combine(outputDir, imageFileName);

                // Save the image (icon) to disk.
                shape.ImageData.Save(imagePath);
                extractedCount++;
            }
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No OLE object images were found to extract.");

        // Validate that at least one image file was created.
        if (!Directory.GetFiles(outputDir, "*.png").Any() &&
            !Directory.GetFiles(outputDir, "*.jpg").Any() &&
            !Directory.GetFiles(outputDir, "*.jpeg").Any() &&
            !Directory.GetFiles(outputDir, "*.bmp").Any())
        {
            throw new InvalidOperationException("Extraction completed but no image files were created.");
        }
    }
}
