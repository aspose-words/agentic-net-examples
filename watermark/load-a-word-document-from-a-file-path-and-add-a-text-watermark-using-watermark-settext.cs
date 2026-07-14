using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample source document.
        string inputPath = Path.Combine(artifactsDir, "input.docx");
        Document sampleDoc = new Document();
        var builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document.");
        sampleDoc.Save(inputPath);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Add a text watermark.
        doc.Watermark.SetText("Confidential");

        // Save the watermarked document.
        string outputPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputPath);
    }
}
