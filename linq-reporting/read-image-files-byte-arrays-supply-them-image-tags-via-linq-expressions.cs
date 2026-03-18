using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageTagProcessor
{
    static void Main()
    {
        // Directory that contains the source images.
        string imageDirectory = @"C:\Images";

        // Load all image files from the directory and create a dictionary
        // where the key is the file name without extension and the value is the image bytes.
        var imageBytesByName = Directory.GetFiles(imageDirectory, "*.*")
            .Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase) ||
                        f.EndsWith(".gif", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(
                f => Path.GetFileNameWithoutExtension(f),
                f => File.ReadAllBytes(f));

        // Load the Word document that contains image tags (e.g., shapes with a title that matches the image name).
        Document doc = new Document(@"C:\Docs\Template.docx");

        // Find all shapes that can contain an image.
        var imageShapes = doc.GetChildNodes(NodeType.Shape, true)
            .Cast<Shape>()
            .Where(s => s.HasImage);

        // Replace each shape's image with the corresponding byte array from the dictionary.
        foreach (Shape shape in imageShapes)
        {
            // Use the shape's Title property as the lookup key.
            // Adjust this to use AlternativeText or any other identifier as needed.
            string? key = shape.Title; // Title may be nullable in newer .NET versions.

            if (string.IsNullOrEmpty(key))
                continue;

            if (imageBytesByName.TryGetValue(key, out byte[] bytes))
            {
                // Set the image from the byte array via a MemoryStream.
                // Aspose.Words.ImageData.SetImage does not accept a raw byte[] in recent versions;
                // it expects a Stream, a file path, an Image, or an SKBitmap.
                using var ms = new MemoryStream(bytes);
                shape.ImageData.SetImage(ms);
            }
        }

        // Save the modified document.
        doc.Save(@"C:\Docs\Result.docx");
    }
}
