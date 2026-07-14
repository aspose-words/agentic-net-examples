using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, Pen

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare deterministic folders.
        // -----------------------------------------------------------------
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        string extractedDir = Path.Combine(artifactsDir, "Extracted");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(extractedDir);

        // -----------------------------------------------------------------
        // 2. Create a sample image (sample.png) using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 200;
        const int imgHeight = 200;

        Bitmap bitmap = new Bitmap(imgWidth, imgHeight);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.White);
        Pen pen = new Pen(Color.Blue, 5);
        graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // -----------------------------------------------------------------
        // 3. Create a Word document and insert the sample image.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 4. Load the document and extract all embedded images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        List<string> extractedImagePaths = new List<string>();
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string extractedImagePath = Path.Combine(extractedDir, $"extracted_{imageIndex}{extension}");
                shape.ImageData.Save(extractedImagePath);
                extractedImagePaths.Add(extractedImagePath);
                imageIndex++;
            }
        }

        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // -----------------------------------------------------------------
        // 5. Generate a PowerShell script that re‑embeds the extracted images.
        // -----------------------------------------------------------------
        string psScriptPath = Path.Combine(artifactsDir, "reembed_images.ps1");
        using (StreamWriter writer = new StreamWriter(psScriptPath, false))
        {
            writer.WriteLine("# PowerShell script to re‑embed extracted images into a new Word document");
            writer.WriteLine("$word = New-Object -ComObject Word.Application");
            writer.WriteLine("$word.Visible = $false");
            writer.WriteLine("$doc = $word.Documents.Add()");
            writer.WriteLine("$range = $doc.Content");

            foreach (string imgPath in extractedImagePaths)
            {
                // Escape backslashes for PowerShell string literals.
                string psImgPath = imgPath.Replace("\\", "\\\\");
                writer.WriteLine($"$range.InlineShapes.AddPicture(\"{psImgPath}\")");
                writer.WriteLine("$range.InsertParagraphAfter()");
            }

            string reembeddedDocPath = Path.Combine(artifactsDir, "reembedded.docx").Replace("\\", "\\\\");
            writer.WriteLine($"$doc.SaveAs([ref]\"{reembeddedDocPath}\")");
            writer.WriteLine("$doc.Close()");
            writer.WriteLine("$word.Quit()");
        }

        // -----------------------------------------------------------------
        // 6. Validation – ensure the PowerShell script was created.
        // -----------------------------------------------------------------
        if (!File.Exists(psScriptPath))
            throw new FileNotFoundException("PowerShell script was not created.", psScriptPath);
    }
}
