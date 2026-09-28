using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample content for DOCX to XLSX conversion.");
        string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the created DOCX document.
        Document doc = new Document(inputPath);

        // Configure XLSX save options with maximum compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum
        };

        // Save the document as XLSX using the configured options.
        string outputPath = "output.xlsx";
        doc.Save(outputPath, xlsxOptions);

        // Validate that the XLSX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }

        // Optional cleanup (comment out if you want to inspect the files).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
