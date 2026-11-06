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
        builder.Writeln("Sample content for DOCX to XLSX conversion.");
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the created DOCX document.
        Document doc = new Document(inputPath);

        // Configure XlsxSaveOptions with maximum compression.
        XlsxSaveOptions saveOptions = new XlsxSaveOptions
        {
            // Use the Aspose.Words.Saving.CompressionLevel enum to avoid ambiguity.
            CompressionLevel = Aspose.Words.Saving.CompressionLevel.Maximum
        };

        // Save the document as XLSX using the configured options.
        const string outputPath = "output.xlsx";
        doc.Save(outputPath, saveOptions);

        // Validate that the XLSX file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }

        // Optional: Clean up the temporary input file.
        File.Delete(inputPath);
    }
}
