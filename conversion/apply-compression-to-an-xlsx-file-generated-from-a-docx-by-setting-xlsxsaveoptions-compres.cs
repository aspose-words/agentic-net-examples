using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the intermediate DOCX and final XLSX files
        string inputDocxPath = Path.Combine(artifactsDir, "input.docx");
        string outputXlsxPath = Path.Combine(artifactsDir, "output.xlsx");

        // Create a simple DOCX document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for XLSX conversion.");
        doc.Save(inputDocxPath, SaveFormat.Docx);

        // Load the DOCX document
        Document loadedDoc = new Document(inputDocxPath);

        // Set up XLSX save options with fast compression
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the configured options
        loadedDoc.Save(outputXlsxPath, xlsxOptions);

        // Verify that the XLSX file was created
        if (!File.Exists(outputXlsxPath))
        {
            throw new InvalidOperationException("Expected XLSX output was not created.");
        }
    }
}
