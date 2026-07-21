using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names for the input DOC and the output XLSX.
        const string inputPath = "sample.doc";
        const string outputPath = "converted.xlsx";

        // -----------------------------------------------------------------
        // 1. Create a sample DOC file.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOC content for conversion to XLSX.");
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // 2. Load the DOC file that was just created.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Prepare XlsxSaveOptions with the default compression level.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Normal, // default compression
            SaveFormat = SaveFormat.Xlsx               // must be set for XlsxSaveOptions
        };

        // -----------------------------------------------------------------
        // 4. Save the document as an XLSX workbook using the options.
        // -----------------------------------------------------------------
        doc.Save(outputPath, xlsxOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the XLSX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");

        // Optional: inform that the conversion succeeded.
        Console.WriteLine($"Conversion completed successfully. Output file: {Path.GetFullPath(outputPath)}");
    }
}
