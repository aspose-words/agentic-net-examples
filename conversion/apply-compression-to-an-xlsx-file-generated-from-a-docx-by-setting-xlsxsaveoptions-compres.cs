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
        builder.Writeln("This is a sample document for XLSX conversion.");
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document loadedDoc = new Document(inputPath);

        // Set up XlsxSaveOptions with Fast compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the specified options.
        const string outputPath = "output.xlsx";
        loadedDoc.Save(outputPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("XLSX file was not created.");

        // Output the size of the generated file.
        FileInfo fileInfo = new FileInfo(outputPath);
        Console.WriteLine($"XLSX saved with Fast compression. Size: {fileInfo.Length} bytes.");
    }
}
