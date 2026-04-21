using System;
using System.IO;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare directories.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(baseDir);

        // Create a simple source document.
        string sourcePath = Path.Combine(baseDir, "source.docx");
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document.");
        sourceDoc.Save(sourcePath);

        // Simulate a network share by copying the file to a subfolder.
        string networkShareDir = Path.Combine(baseDir, "NetworkShare");
        Directory.CreateDirectory(networkShareDir);
        string networkDocPath = Path.Combine(networkShareDir, "sample.docx");
        File.Copy(sourcePath, networkDocPath, true);

        // Open the document from the simulated network share using a FileStream.
        // The using statement guarantees that the file handle is released.
        using (FileStream stream = File.Open(networkDocPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            var doc = new Document(stream);

            // Configure text watermark options.
            var watermarkOptions = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 48,
                Color = Color.LightGray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Add the watermark.
            doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

            // Save the watermarked document to an output file.
            string outputPath = Path.Combine(baseDir, "watermarked.docx");
            doc.Save(outputPath);
        }

        // Verify that the output file was created.
        if (File.Exists(Path.Combine(baseDir, "watermarked.docx")))
        {
            Console.WriteLine("Watermark applied and document saved successfully.");
        }
    }
}
