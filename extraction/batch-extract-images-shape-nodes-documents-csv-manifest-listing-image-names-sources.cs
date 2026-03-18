using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class ImageExtractor
{
    static void Main()
    {
        // Path to the source Word document.
        string sourceDocPath = @"C:\Docs\InputDocument.docx";

        // Directory where extracted images and the CSV manifest will be saved.
        string outputDir = @"C:\Docs\ExtractedImages";
        Directory.CreateDirectory(outputDir);

        // Load the document using Aspose.Words.
        Document doc = new Document(sourceDocPath);

        // Collect all Shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Prepare a StringBuilder for the CSV manifest.
        // Columns: ImageFileName, ShapeName, ShapeTitle, SourceFullName (if linked), ImageType
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("ImageFileName,ShapeName,ShapeTitle,SourceFullName,ImageType");

        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (!shape.HasImage)
                continue;

            // Determine a file extension based on the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

            // Build a unique image file name.
            string imageFileName = $"Image_{imageIndex}{extension}";
            string imagePath = Path.Combine(outputDir, imageFileName);

            // Save the image to the file system.
            shape.ImageData.Save(imagePath);

            // Gather optional metadata for the CSV.
            string shapeName = shape.Name ?? string.Empty;
            string shapeTitle = shape.Title ?? string.Empty;
            string sourceFullName = shape.ImageData.SourceFullName ?? string.Empty;
            string imageType = shape.ImageData.ImageType.ToString();

            // Append a line to the CSV manifest.
            csvBuilder.AppendLine($"{imageFileName},{EscapeCsv(shapeName)},{EscapeCsv(shapeTitle)},{EscapeCsv(sourceFullName)},{imageType}");

            imageIndex++;
        }

        // Write the CSV manifest to a file named "ImageManifest.csv" in the output directory.
        string csvPath = Path.Combine(outputDir, "ImageManifest.csv");
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }

    // Helper method to escape CSV fields that may contain commas or quotes.
    private static string EscapeCsv(string field)
    {
        if (field.Contains(',') || field.Contains('\"') || field.Contains('\n') || field.Contains('\r'))
        {
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
