using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string imageFile = "input.png";
        const string docFile = "sample.docx";
        const string scriptFile = "reembed.ps1";

        // -------------------------------------------------
        // 1. Create a deterministic sample image.
        // -------------------------------------------------
        const int imgWidth = 100;
        const int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
            }
            bitmap.Save(imageFile, ImageFormat.Png);
        }

        // Verify the image was created.
        if (!File.Exists(imageFile))
            throw new FileNotFoundException($"Failed to create sample image '{imageFile}'.");

        // -------------------------------------------------
        // 2. Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imageFile);
        doc.Save(docFile);

        // Verify the document was saved.
        if (!File.Exists(docFile))
            throw new FileNotFoundException($"Failed to save document '{docFile}'.");

        // -------------------------------------------------
        // 3. Load the document and extract all images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docFile);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedName = $"extracted_{extractedCount}{extension}";
                shape.ImageData.Save(extractedName);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -------------------------------------------------
        // 4. Generate a PowerShell script that can re‑embed the extracted images.
        // -------------------------------------------------
        using (StreamWriter writer = new StreamWriter(scriptFile, false))
        {
            writer.WriteLine("# PowerShell script to re‑embed extracted images into a Word document");
            writer.WriteLine("# This script assumes Aspose.Words for .NET is available to the PowerShell runtime.");
            writer.WriteLine();
            writer.WriteLine("$docPath = \"{0}\"", docFile);
            writer.WriteLine("$outputPath = \"reembedded.docx\"");
            writer.WriteLine();
            writer.WriteLine("Add-Type -Path \"Aspose.Words.dll\"");
            writer.WriteLine("Add-Type -Path \"Aspose.Drawing.dll\"");
            writer.WriteLine();
            writer.WriteLine("$doc = New-Object Aspose.Words.Document $docPath");
            writer.WriteLine("$builder = New-Object Aspose.Words.DocumentBuilder $doc");
            writer.WriteLine();
            for (int i = 0; i < extractedCount; i++)
            {
                string imgName = $"extracted_{i}{FileFormatUtil.ImageTypeToExtension(loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>().ElementAt(i).ImageData.ImageType)}";
                writer.WriteLine("$builder.Writeln(\"Inserting image: {0}\")", imgName);
                writer.WriteLine("$builder.InsertImage(\"{0}\")", imgName);
                writer.WriteLine("$builder.Writeln()");
            }
            writer.WriteLine();
            writer.WriteLine("$doc.Save($outputPath)");
            writer.WriteLine("Write-Host \"Re‑embedded document saved to $outputPath\"");
        }

        // Verify the script file was created.
        if (!File.Exists(scriptFile))
            throw new FileNotFoundException($"Failed to create PowerShell script '{scriptFile}'.");

        // -------------------------------------------------
        // 5. Clean up temporary resources (optional).
        // -------------------------------------------------
        // No explicit cleanup required; files remain for inspection.
    }
}
