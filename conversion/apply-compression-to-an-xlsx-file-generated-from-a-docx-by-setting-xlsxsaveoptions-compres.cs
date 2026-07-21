using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary DOCX input and the final XLSX output.
        string inputPath = "input.docx";
        string outputPath = "output.xlsx";

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for XLSX conversion.");

        // Save the DOCX to disk (bootstrap the input file).
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document from the file system.
        Document loadedDoc = new Document(inputPath);

        // Set up XlsxSaveOptions with fast compression.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        xlsxOptions.CompressionLevel = CompressionLevel.Fast;
        xlsxOptions.SaveFormat = SaveFormat.Xlsx;

        // Save the document as XLSX using the specified options.
        loadedDoc.Save(outputPath, xlsxOptions);

        // Verify that the XLSX file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected XLSX output was not created.");
    }
}
