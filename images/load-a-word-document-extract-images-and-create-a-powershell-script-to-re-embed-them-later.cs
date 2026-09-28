using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image.
        const string sampleImagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
            }
            bitmap.Save(sampleImagePath);
        }

        // Create a Word document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document and extract all embedded images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string extractedImagePath = $"extracted-{extractedCount}.png";
                shape.ImageData.Save(extractedImagePath);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Generate a PowerShell script that re‑embeds the extracted images.
        const string psScriptPath = "reembed.ps1";
        string psScriptContent = @"
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open((Resolve-Path 'sample.docx').Path)
foreach ($img in Get-ChildItem -Path . -Filter 'extracted-*.png') {
    $range = $doc.Content
    $range.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
    $range.InlineShapes.AddPicture($img.FullName) | Out-Null
}
$doc.Save()
$doc.Close()
$word.Quit()
".TrimStart('\r', '\n');
        File.WriteAllText(psScriptPath, psScriptContent);

        // Validate that the PowerShell script was created.
        if (!File.Exists(psScriptPath))
        {
            throw new InvalidOperationException("Failed to create the PowerShell script.");
        }
    }
}
