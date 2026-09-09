using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content.");
        const string inputPath = "input.doc";
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Configure XLSX save options with default compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // CompressionLevel defaults to Normal; set explicitly for clarity.
            CompressionLevel = CompressionLevel.Normal,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save as XLSX.
        const string outputPath = "output.xlsx";
        doc.Save(outputPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected XLSX output file was not created.");

        // Clean up temporary files (optional).
        File.Delete(inputPath);
    }
}
