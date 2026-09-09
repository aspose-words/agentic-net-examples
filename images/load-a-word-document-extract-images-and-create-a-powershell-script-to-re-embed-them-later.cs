using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names and folders.
        string workDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(workDir, "sample.png");
        string docPath = Path.Combine(workDir, "sample.docx");
        string scriptPath = Path.Combine(workDir, "reembed_images.ps1");

        // -------------------------------------------------
        // 1. Create a deterministic sample image (PNG).
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;
        var bitmap = new Bitmap(imgWidth, imgHeight);
        var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Draw a simple rectangle to make the image non‑empty.
        graphics.DrawRectangle(new Pen(Color.Black, 2), 10, 10, imgWidth - 20, imgHeight - 20);
        bitmap.Save(imagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -------------------------------------------------
        // 2. Create a Word document and insert the image twice.
        // -------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with embedded images:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all images.
        // -------------------------------------------------
        var loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedImagePath = Path.Combine(workDir, $"extracted_{imageIndex}{extension}");
                shape.ImageData.Save(extractedImagePath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Generate a PowerShell script that re‑embeds the extracted images.
        // -------------------------------------------------
        using (var writer = new StreamWriter(scriptPath, false))
        {
            writer.WriteLine("# PowerShell script to re‑embed extracted images into a Word document");
            writer.WriteLine("$assemblyPath = \"{0}\"" , typeof(Document).Assembly.Location.Replace("\\", "/"));
            writer.WriteLine("Add-Type -Path $assemblyPath");
            writer.WriteLine("$doc = New-Object Aspose.Words.Document \"{0}\"" , docPath.Replace("\\", "/"));
            writer.WriteLine("$builder = New-Object Aspose.Words.DocumentBuilder $doc");
            for (int i = 0; i < imageIndex; i++)
            {
                string ext = Path.GetExtension($"extracted_{i}"); // extension already includes dot
                string imgFile = Path.Combine(workDir, $"extracted_{i}{ext}").Replace("\\", "/");
                writer.WriteLine("$builder.InsertImage(\"{0}\")", imgFile);
            }
            writer.WriteLine("$doc.Save(\"{0}\")", Path.Combine(workDir, "reembedded.docx").Replace("\\", "/"));
        }

        // Validate that the PowerShell script was created.
        if (!File.Exists(scriptPath))
            throw new InvalidOperationException("Failed to create the PowerShell script.");

        // -------------------------------------------------
        // 5. Clean up temporary resources (optional).
        // -------------------------------------------------
        // All disposable objects have been disposed via using statements or explicit Dispose calls.
    }
}
