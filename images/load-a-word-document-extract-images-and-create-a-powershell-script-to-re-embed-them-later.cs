using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class ImageExtractionExample
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample image (sample.png).
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Build a sample Word document that contains the image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateSampleDocument(docPath, sampleImagePath);

        // 3. Load the document and extract all embedded images.
        List<string> extractedImages = ExtractImagesFromDocument(docPath, artifactsDir);

        // Validate that at least one image was extracted.
        if (extractedImages.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 4. Write a JSON file that lists the extracted images (demonstrates Newtonsoft.Json usage).
        string jsonPath = Path.Combine(artifactsDir, "extracted_images.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(extractedImages, Formatting.Indented));

        // 5. Generate a PowerShell script that can re‑embed the extracted images into the document.
        string psScriptPath = Path.Combine(artifactsDir, "reembed_images.ps1");
        GeneratePowerShellScript(psScriptPath, docPath, extractedImages);

        // All work is done; the program exits automatically.
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a bitmap, fill it with white, and draw a simple rectangle.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a blue rectangle for visual distinction.
            using (Pen pen = new Pen(Color.Blue, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }

            // Save the bitmap to the specified file.
            bitmap.Save(filePath);
        }
    }

    private static void CreateSampleDocument(string docPath, string imagePath)
    {
        // Create a blank document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample document with an embedded image:");
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }

    private static List<string> ExtractImagesFromDocument(string docPath, string outputDir)
    {
        // Load the document.
        Document doc = new Document(docPath);

        // Collect all shapes that contain images.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        List<string> savedImagePaths = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"extracted_{imageIndex}{extension}";
                string fullPath = Path.Combine(outputDir, imageFileName);

                // Save the image.
                shape.ImageData.Save(fullPath);
                savedImagePaths.Add(fullPath);
                imageIndex++;
            }
        }

        return savedImagePaths;
    }

    private static void GeneratePowerShellScript(string scriptPath, string docPath, List<string> imagePaths)
    {
        // Build the PowerShell script content.
        StringBuilder sb = new StringBuilder();

        sb.AppendLine("$word = New-Object -ComObject Word.Application");
        sb.AppendLine("$word.Visible = $false");
        sb.AppendLine();

        // Use full paths to avoid ambiguity.
        string docFullPath = Path.GetFullPath(docPath).Replace("\\", "\\\\");
        sb.AppendLine($"$doc = $word.Documents.Open(\"{docFullPath}\")");
        sb.AppendLine("$selection = $word.Selection");
        sb.AppendLine();

        // Insert each extracted image at the end of the document.
        foreach (string imgPath in imagePaths)
        {
            string imgFullPath = Path.GetFullPath(imgPath).Replace("\\", "\\\\");
            sb.AppendLine($"$selection.EndKey([Microsoft.Office.Interop.Word.WdUnits]::wdStory)");
            sb.AppendLine($"$selection.InlineShapes.AddPicture(\"{imgFullPath}\")");
            sb.AppendLine();
        }

        sb.AppendLine("$doc.Save()");
        sb.AppendLine("$doc.Close()");
        sb.AppendLine("$word.Quit()");
        sb.AppendLine();

        // Write the script to file.
        File.WriteAllText(scriptPath, sb.ToString());
    }
}
