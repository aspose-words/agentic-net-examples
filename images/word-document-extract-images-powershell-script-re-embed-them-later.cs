using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class ImageExtractor
{
    static void Main()
    {
        // Path to the source Word document
        string sourceDocPath = "input.docx";

        // Directory where extracted images will be saved
        string imagesDir = "ExtractedImages";
        Directory.CreateDirectory(imagesDir);

        // Load the document (uses the Document(string) constructor rule)
        Document doc = new Document(sourceDocPath);

        // Collect all Shape nodes (including images) from the document
        var shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the full file name for the extracted image
                string imageFilePath = Path.Combine(imagesDir, $"image_{imageIndex}{extension}");

                // Save the image data to the file system
                shape.ImageData.Save(imageFilePath);

                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // Generate a PowerShell script that will re‑embed the extracted images
        // -----------------------------------------------------------------
        string psScriptPath = "ReembedImages.ps1";

        using (StreamWriter writer = new StreamWriter(psScriptPath, false))
        {
            writer.WriteLine("# PowerShell script to re‑embed extracted images into a Word document");
            writer.WriteLine("$sourceDoc = \"input.docx\"");
            writer.WriteLine("$outputDoc = \"output.docx\"");
            writer.WriteLine("");
            writer.WriteLine("# Load Aspose.Words assembly (adjust the path if necessary)");
            writer.WriteLine("[Reflection.Assembly]::LoadFile((Resolve-Path \"{0}\").Path) | Out-Null", typeof(Document).Assembly.Location);
            writer.WriteLine("");
            writer.WriteLine("# Create Document and DocumentBuilder objects");
            writer.WriteLine("$doc = New-Object Aspose.Words.Document $sourceDoc");
            writer.WriteLine("$builder = New-Object Aspose.Words.DocumentBuilder $doc");
            writer.WriteLine("");

            // Insert each extracted image back into the document
            for (int i = 0; i < imageIndex; i++)
            {
                // Use a wildcard to locate the image file regardless of its extension
                writer.WriteLine("$imgPath = Get-ChildItem -Path \"{0}\" -File | Where-Object {{ $_.Name -like \"image_{1}*\" }} | Select-Object -First 1", imagesDir, i);
                writer.WriteLine("if ($imgPath) {{ $builder.InsertImage($imgPath.FullName) }}");
            }

            writer.WriteLine("");
            writer.WriteLine("# Save the modified document");
            writer.WriteLine("$doc.Save($outputDoc)");
        }

        Console.WriteLine("Image extraction complete.");
        Console.WriteLine($"Extracted images saved to: {Path.GetFullPath(imagesDir)}");
        Console.WriteLine($"PowerShell script generated at: {Path.GetFullPath(psScriptPath)}");
    }
}
