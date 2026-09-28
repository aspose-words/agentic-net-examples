using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document to be converted to XLSX.");
        string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure XLSX save options with fast compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast
        };

        // Save the document as XLSX using the configured options.
        string outputPath = "output.xlsx";
        doc.Save(outputPath, xlsxOptions);

        // Validate that the XLSX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The XLSX file was not created as expected.");
        }

        // Optional cleanup (comment out if you want to inspect the files).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
