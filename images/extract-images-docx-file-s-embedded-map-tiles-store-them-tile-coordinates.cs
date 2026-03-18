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
            // Path to the source DOCX that contains embedded map tiles.
            string sourceDocPath = @"C:\Input\MapDocument.docx";

            // Folder where extracted tile images will be saved.
            string outputFolder = @"C:\Output\Tiles";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the document (uses the provided Document constructor rule).
            Document doc = new Document(sourceDocPath);

            // Get all Shape nodes in the document (including those inside headers/footers).
            NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapeNodes.OfType<Shape>())
            {
                // Process only shapes that actually contain an image.
                if (!shape.HasImage)
                    continue;

                // Attempt to obtain tile coordinates from the shape's AlternativeText.
                // Expected format: "X=12;Y=34" or "12_34". Adjust parsing as needed.
                string altText = shape.AlternativeText?.Trim();
                if (string.IsNullOrEmpty(altText))
                    continue; // No coordinate information; skip.

                // Parse coordinates.
                int tileX = 0, tileY = 0;
                bool parsed = TryParseCoordinates(altText, out tileX, out tileY);
                if (!parsed)
                    continue; // Unable to parse coordinates; skip.

                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name using tile coordinates.
                string fileName = $"tile_{tileX}_{tileY}{extension}";
                string fullPath = Path.Combine(outputFolder, fileName);

                // Save the image data to the file system (uses the provided ImageData.Save rule).
                shape.ImageData.Save(fullPath);
            }
        }

        /// <summary>
        /// Parses a coordinate string and extracts X and Y integer values.
        /// Supports formats like "12_34", "X=12;Y=34", "12,34", etc.
        /// </summary>
        private static bool TryParseCoordinates(string text, out int x, out int y)
        {
            x = y = 0;
            // Remove any whitespace.
            text = text.Replace(" ", string.Empty).ToUpperInvariant();

            // Split on common delimiters.
            char[] delimiters = new[] { '_', ',', ';' };
            string[] parts = text.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 2 && int.TryParse(parts[0], out x) && int.TryParse(parts[1], out y))
                return true;

            // Handle "X=12;Y=34" style.
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
