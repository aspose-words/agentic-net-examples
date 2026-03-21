using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ExtractImagesBySection
{
    static void Main()
    {
        // Directory where extracted images will be saved.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedImages");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with one section and an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with an image:");

        // Add an image from a base64 string (a tiny red dot PNG).
        const string base64Png = 
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAusB9YkK6V8AAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            builder.InsertImage(ms);
        }

        // Iterate through each section in the document.
        for (int secIdx = 0; secIdx < doc.Sections.Count; secIdx++)
        {
            Section section = doc.Sections[secIdx];
            NodeCollection shapes = section.GetChildNodes(NodeType.Shape, true);
            int imageIdx = 0;

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                    string fileName = $"Section_{secIdx + 1}_Image_{++imageIdx}{extension}";
                    string fullPath = Path.Combine(outputDir, fileName);
                    shape.ImageData.Save(fullPath);
                }
            }
        }

        // Save the (unchanged) document to the output directory.
        string outputDocPath = Path.Combine(outputDir, "ProcessedDocument.docx");
        doc.Save(outputDocPath);
    }
}
