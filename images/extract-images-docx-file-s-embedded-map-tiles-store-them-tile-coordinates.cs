using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace MapTileExtractor
{
    class Program
    {
        static void Main()
        {
            // Folder where extracted tile images will be saved.
            string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Tiles");

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Create a sample document with an embedded image shape that has coordinate info.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A 1x1 pixel PNG (transparent) encoded in base64.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6UAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            using var imageStream = new MemoryStream(pngBytes);
            Shape shape = builder.InsertImage(imageStream);
            shape.AlternativeText = "12_34"; // Example tile coordinates.

            // Get all Shape nodes in the document (including those inside headers/footers).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape s in shapeNodes.OfType<Shape>())
            {
                if (!s.HasImage)
                    continue;

                string altText = s.AlternativeText?.Trim();
                if (string.IsNullOrEmpty(altText))
                    continue;

                if (!TryParseCoordinates(altText, out int tileX, out int tileY))
                    continue;

                string extension = FileFormatUtil.ImageTypeToExtension(s.ImageData.ImageType);
                string fileName = $"tile_{tileX}_{tileY}{extension}";
                string fullPath = Path.Combine(outputFolder, fileName);

                s.ImageData.Save(fullPath);
                Console.WriteLine($"Saved tile image to: {fullPath}");
            }
        }

        private static bool TryParseCoordinates(string text, out int x, out int y)
        {
            x = y = 0;
            text = text.Replace(" ", string.Empty).ToUpperInvariant();

            char[] delimiters = new[] { '_', ',', ';' };
            string[] parts = text.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 2 && int.TryParse(parts[0], out x) && int.TryParse(parts[1], out y))
                return true;

            if (text.Contains("X=") && text.Contains("Y="))
            {
                string[] tokens = text.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string token in tokens)
                {
                    if (token.StartsWith("X=") && int.TryParse(token.Substring(2), out x))
                        continue;
                    if (token.StartsWith("Y=") && int.TryParse(token.Substring(2), out y))
                        continue;
                }
                return x != 0 || y != 0;
            }

            return false;
        }
    }
}
