using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (input.png)
        const string sampleImagePath = "input.png";
        const int imgWidth = 200;
        const int imgHeight = 200;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Additional drawing can be added here if needed
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Ensure the sample image was created
        if (!File.Exists(sampleImagePath))
            throw new FileNotFoundException("Failed to create the sample image.", sampleImagePath);

        // Create a Word document and insert the sample image
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Ensure the document was saved
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create the sample document.", docPath);

        // Load the document and extract all embedded images
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

        // Validate that at least one image was extracted
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Generate a PowerShell script that re‑embeds the extracted images
        StringBuilder psBuilder = new StringBuilder();

        // Note: Adjust the path to Aspose.Words.dll as needed when the script is used.
        string asposeDllPath = Path.GetFullPath("Aspose.Words.dll");
        psBuilder.AppendLine($"Add-Type -Path \"{asposeDllPath}\"");
        psBuilder.AppendLine($"$doc = New-Object Aspose.Words.Document \"{docPath}\"");
        psBuilder.AppendLine("$builder = New-Object Aspose.Words.DocumentBuilder($doc)");
        psBuilder.AppendLine("$builder.MoveToDocumentEnd()");

        for (int i = 0; i < extractedCount; i++)
        {
            string imgPath = $"extracted-{i}.png";
            psBuilder.AppendLine($"$builder.InsertImage(\"{imgPath}\")");
        }

        const string reembeddedDocPath = "reembedded.docx";
        psBuilder.AppendLine($"$doc.Save(\"{reembeddedDocPath}\")");

        const string psScriptPath = "reembed.ps1";
        File.WriteAllText(psScriptPath, psBuilder.ToString());

        // Validate that the PowerShell script was created
        if (!File.Exists(psScriptPath))
            throw new FileNotFoundException("Failed to create the PowerShell script.", psScriptPath);
    }
}
