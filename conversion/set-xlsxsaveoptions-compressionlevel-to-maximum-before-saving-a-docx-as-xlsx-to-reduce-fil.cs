using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string inputPath = "input.docx";
        string outputPath = "output.xlsx";

        // Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for compression test.");
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from the file system.
        Document loadedDoc = new Document(inputPath);

        // Configure XlsxSaveOptions with maximum compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the configured options.
        loadedDoc.Save(outputPath, xlsxOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XLSX file was not created.");

        // Optional: indicate success.
        Console.WriteLine("Document successfully saved as XLSX with maximum compression.");
    }
}
