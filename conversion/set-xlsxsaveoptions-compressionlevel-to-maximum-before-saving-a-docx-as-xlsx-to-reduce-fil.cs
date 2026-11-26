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
        builder.Writeln("This is a sample document to be saved as XLSX.");
        builder.Writeln("It contains multiple lines of text to increase size.");

        // Save the document as DOCX (bootstrap input file).
        string inputPath = "sample.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure XLSX save options with maximum compression.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum
        };

        // Save the document as XLSX using the configured options.
        string outputPath = "output.xlsx";
        doc.Save(outputPath, saveOptions);

        // Validate that the XLSX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }

        // Optionally, verify that the file size is greater than zero.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The generated XLSX file is empty.");
        }

        // Clean up created files (optional).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
