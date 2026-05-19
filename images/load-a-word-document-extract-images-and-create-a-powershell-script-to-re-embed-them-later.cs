using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a deterministic sample image (input.png)
        string sampleImagePath = Path.Combine(outputDir, "input.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create a Word document and insert the sample image
        string docPath = Path.Combine(outputDir, "DocumentWithImages.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Load the document and extract all images
        List<string> extractedImagePaths = ExtractImagesFromDocument(docPath, outputDir);

        // Validate that at least one image was extracted
        if (extractedImagePaths.Count == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // 4. Generate a PowerShell script that can re‑embed the extracted images
        string psScriptPath = Path.Combine(outputDir, "ReembedImages.ps1");
        GeneratePowerShellScript(psScriptPath, extractedImagePaths, outputDir);

        // Validate that the script file was created
        if (!File.Exists(psScriptPath))
            throw new FileNotFoundException("PowerShell script was not created.", psScriptPath);
    }

    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Create a bitmap, clear it to white, and save it as PNG
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        // Optionally draw something deterministic (a black rectangle)
        graphics.DrawRectangle(new Pen(Color.Black), 10, 10, width - 20, height - 20);
        bitmap.Save(filePath);
        graphics.Dispose();
        bitmap.Dispose();
    }

    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image into the document
        Shape shape = builder.InsertImage(imagePath);
        // Ensure the shape is appended (InsertImage already does this)
        doc.Save(docPath);
    }

    private static List<string> ExtractImagesFromDocument(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedPaths = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string imageFileName = $"extracted_image_{imageIndex}{extension}";
                string imageFullPath = Path.Combine(outputDir, imageFileName);
                shape.ImageData.Save(imageFullPath);
                extractedPaths.Add(imageFullPath);
                imageIndex++;
            }
        }

        return extractedPaths;
    }

    private static void GeneratePowerShellScript(string scriptPath, List<string> imagePaths, string outputDir)
    {
        // Build the PowerShell script content
        var lines = new List<string>
        {
            "# PowerShell script to re‑embed extracted images into a new Word document",
            "Add-Type -Path \"Aspose.Words.dll\"",
            "",
            "$doc = New-Object Aspose.Words.Document()",
            "$builder = New-Object Aspose.Words.DocumentBuilder($doc)",
            "",
            "$builder.Writeln(\"Re‑embedded images:\")",
            ""
        };

        // Add the image list
        lines.Add("$images = @(");
        for (int i = 0; i < imagePaths.Count; i++)
        {
            string escapedPath = imagePaths[i].Replace("\\", "\\\\");
            string line = $"    \"{escapedPath}\"{(i < imagePaths.Count - 1 ? "," : "")}";
            lines.Add(line);
        }
        lines.Add(")");
        lines.Add("");
        lines.Add("foreach ($img in $images) {");
        lines.Add("    $builder.InsertImage($img)");
        lines.Add("}");
        lines.Add("");
        string outputDocPath = Path.Combine(outputDir, "ReembeddedDocument.docx").Replace("\\", "\\\\");
        lines.Add($"$doc.Save(\"{outputDocPath}\")");
        lines.Add("");

        File.WriteAllLines(scriptPath, lines);
    }
}
