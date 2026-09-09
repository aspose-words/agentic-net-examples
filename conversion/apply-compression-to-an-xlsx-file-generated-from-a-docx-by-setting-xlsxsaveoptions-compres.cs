using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the intermediate DOCX and final XLSX files.
        const string docxPath = "sample.docx";
        const string xlsxPath = "compressed.xlsx";

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for XLSX conversion.");
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document we just created.
        Document loadedDoc = new Document(docxPath);

        // Set up XLSX save options with fast compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the specified options.
        loadedDoc.Save(xlsxPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("The XLSX file was not created.");

        Console.WriteLine($"XLSX file saved with Fast compression: {xlsxPath}");
    }
}
