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
        builder.Writeln("Sample content for DOC to XLSX conversion.");
        string inputPath = "sample.doc";
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Prepare XlsxSaveOptions with default compression (Normal).
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Normal
        };

        // Save as XLSX.
        string outputPath = "output.xlsx";
        doc.Save(outputPath, saveOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }
    }
}
