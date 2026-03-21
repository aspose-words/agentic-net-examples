using System;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageExtractor
{
    static void Main()
    {
        // Use paths relative to the executable directory.
        string baseDir = AppContext.BaseDirectory;
        string sourceDocPath = Path.Combine(baseDir, "InputDocument.docx");
        string outputDir = Path.Combine(baseDir, "ExtractedImages");
        Directory.CreateDirectory(outputDir);

        // Create a minimal document if the source file does not exist.
        if (!File.Exists(sourceDocPath))
        {
            var tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            builder.Writeln("Sample document created because the original file was missing.");
            tempDoc.Save(sourceDocPath);
        }

        // Load the document.
        Document doc = new Document(sourceDocPath);

        // Collect all Shape nodes (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        var csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("ImageFileName,ShapeName,ShapeTitle,SourceFullName,ImageType");

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string imageFileName = $"Image_{imageIndex}{extension}";
            string imagePath = Path.Combine(outputDir, imageFileName);
            shape.ImageData.Save(imagePath);

            string shapeName = shape.Name ?? string.Empty;
            string shapeTitle = shape.Title ?? string.Empty;
            string sourceFullName = shape.ImageData.SourceFullName ?? string.Empty;
            string imageType = shape.ImageData.ImageType.ToString();

            csvBuilder.AppendLine($"{imageFileName},{EscapeCsv(shapeName)},{EscapeCsv(shapeTitle)},{EscapeCsv(sourceFullName)},{imageType}");
            imageIndex++;
        }

        string csvPath = Path.Combine(outputDir, "ImageManifest.csv");
        File.WriteAllText(csvPath, csvBuilder.ToString(), Encoding.UTF8);
    }

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
