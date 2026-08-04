using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, PDF/A-1b with embedded ICC profile!");

        // Configure PDF/A-1b compliance.
        // Aspose.Words automatically embeds the sRGB ICC profile when PDF/A compliance is set.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        string outputPath = "output_pdfa1b.pdf";
        doc.Save(outputPath, saveOptions);

        // Validate that the file was created and is not empty.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The file '{outputPath}' was not created.");
        }

        FileInfo fileInfo = new FileInfo(outputPath);
        if (fileInfo.Length == 0)
        {
            throw new InvalidOperationException("The output PDF file is empty.");
        }

        Console.WriteLine($"PDF/A-1b file saved successfully: {outputPath} ({fileInfo.Length} bytes)");
    }
}
