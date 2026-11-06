using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that will be saved as PDF/A‑3u.");

        // Configure PDF/A‑3u save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA3u
        };

        // Save the document as PDF/A‑3u.
        string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Validate that the file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑3u file was not created.");

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The PDF/A‑3u file is empty.");

        // The file has been successfully created.
        Console.WriteLine($"PDF/A‑3u file created at: {Path.GetFullPath(outputPath)} (size: {info.Length} bytes)");
    }
}
