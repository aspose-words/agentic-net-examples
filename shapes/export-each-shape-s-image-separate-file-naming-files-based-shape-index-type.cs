using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public static class ShapeImageExporter
{
    /// <summary>
    /// Exports the image of each shape in the document to a separate file.
    /// The file name is composed of the shape index, the shape type, and the appropriate image extension.
    /// </summary>
    /// <param name="documentPath">Full path to the source Word document.</param>
    /// <param name="outputFolder">Folder where the extracted images will be saved.</param>
    public static void ExportShapeImages(string documentPath, string outputFolder)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Load the document.
        Document doc = new Document(documentPath);

        // Get all Shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int shapeIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain image data.
            if (shape.HasImage)
            {
                // Determine the file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build a unique file name: shape_{index}_{ShapeType}{extension}
                string fileName = $"shape_{shapeIndex}_{shape.ShapeType}{extension}";
                string filePath = Path.Combine(outputFolder, fileName);

                // Save the image data to the file system.
                shape.ImageData.Save(filePath);

                shapeIndex++;
            }
        }
    }
}

public class Program
{
    /// <summary>
    /// Entry point required for a console application.
    /// Adjust the arguments or hard‑code paths for testing.
    /// </summary>
    public static void Main(string[] args)
    {
        // Simple argument handling: first argument = document path, second = output folder.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ShapeImageExporter <documentPath> <outputFolder>");
            return;
        }

        string documentPath = args[0];
        string outputFolder = args[1];

        try
        {
            ShapeImageExporter.ExportShapeImages(documentPath, outputFolder);
            Console.WriteLine("Images exported successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
