using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "input.doc";
        const string outputPath = "output.xlsx";

        // -----------------------------------------------------------------
        // Create a sample DOC file.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Load the DOC file.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Prepare XLSX save options with default compression.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Normal, // Default value.
            SaveFormat = SaveFormat.Xlsx
        };

        // -----------------------------------------------------------------
        // Convert and save as XLSX.
        // -----------------------------------------------------------------
        doc.Save(outputPath, xlsxOptions);

        // -----------------------------------------------------------------
        // Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected XLSX output file was not created.");

        // Optional: inform that conversion succeeded.
        Console.WriteLine($"DOC file '{inputPath}' was successfully converted to XLSX file '{outputPath}'.");
    }
}
