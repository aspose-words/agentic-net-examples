using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string inputPath = "sample.docx";
        string outputPath = "compressed.xlsx";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that will be saved as XLSX.");
        // Save the DOCX so it can be loaded later if needed.
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document (demonstrates the load step).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Configure XlsxSaveOptions with fast compression.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast,
            SaveFormat = SaveFormat.Xlsx
        };

        // -----------------------------------------------------------------
        // 4. Save the document as XLSX using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPath, xlsxOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the XLSX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XLSX file was not created.");

        // Optional: output a confirmation message.
        Console.WriteLine($"XLSX file saved successfully with fast compression: {outputPath}");
    }
}
