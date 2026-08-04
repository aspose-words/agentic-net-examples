using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        const string inputPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for XLSX conversion.");
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loadedDoc = new Document(inputPath);

        // Set up XlsxSaveOptions with maximum compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the configured options.
        const string outputPath = "output.xlsx";
        loadedDoc.Save(outputPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }
    }
}
