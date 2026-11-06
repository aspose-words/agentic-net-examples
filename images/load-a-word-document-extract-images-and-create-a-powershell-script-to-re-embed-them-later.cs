using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define deterministic file names
        string workingDir = Directory.GetCurrentDirectory();
        string sampleImagePath = Path.Combine(workingDir, "input.png");
        string docPath = Path.Combine(workingDir, "sample.docx");
        string extractedImagesDir = Path.Combine(workingDir, "extracted_images");
        string psScriptPath = Path.Combine(workingDir, "reembed_images.ps1");

        // Ensure output directories exist
        Directory.CreateDirectory(extractedImagesDir);

        // -------------------------------------------------
        // 1. Create a deterministic sample image (input.png)
        // -------------------------------------------------
        int imgWidth = 200;
        int imgHeight = 100;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle for visual distinction
                using (Pen pen = new Pen(Color.Blue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all embedded images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImagePaths = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedPath = Path.Combine(extractedImagesDir, $"extracted-{imageIndex}.png");
                shape.ImageData.Save(extractedPath);
                extractedImagePaths.Add(extractedPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // -------------------------------------------------
        // 4. Generate a PowerShell script that re‑embeds the images
        // -------------------------------------------------
        // Determine the path to the Aspose.Words assembly for the script
        string asposeWordsAssemblyPath = typeof(Document).Assembly.Location.Replace("\\", "/");

        // Build PowerShell script content
        var psLines = new List<string>
        {
            "# PowerShell script to re‑embed extracted images into a new Word document",
            $"Add-Type -Path \"{asposeWordsAssemblyPath}\"",
            "",
            "$doc = New-Object Aspose.Words.Document",
            "$builder = New-Object Aspose.Words.DocumentBuilder($doc)",
            ""
        };

        foreach (string imgPath in extractedImagePaths)
        {
            string escapedPath = imgPath.Replace("\\", "/");
            psLines.Add($"$builder.InsertImage(\"{escapedPath}\")");
        }

        psLines.Add("");
        psLines.Add("$outputDoc = \"reembedded.docx\"");
        psLines.Add("$doc.Save($outputDoc)");
        psLines.Add("Write-Host \"Re‑embedded document saved to $outputDoc\"");

        // Write the script to file
        File.WriteAllLines(psScriptPath, psLines);

        // Validate that the PowerShell script was created
        if (!File.Exists(psScriptPath))
        {
            throw new InvalidOperationException("Failed to create the PowerShell script.");
        }

        // -------------------------------------------------
        // 5. Output summary (optional, not interactive)
        // -------------------------------------------------
        var summary = new
        {
            SampleImage = sampleImagePath,
            Document = docPath,
            ExtractedImages = extractedImagePaths,
            PowerShellScript = psScriptPath
        };
        string jsonSummary = JsonConvert.SerializeObject(summary, Formatting.Indented);
        Console.WriteLine(jsonSummary);
    }
}
