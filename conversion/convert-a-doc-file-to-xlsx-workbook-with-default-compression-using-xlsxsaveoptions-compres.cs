using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths for the sample DOC input and the XLSX output.
        const string inputDocPath = "Sample.doc";
        const string outputXlsxPath = "Converted.xlsx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOC file.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample document created for conversion.");
        // Save the document in DOC format (binary Word 97‑2003).
        sampleDoc.Save(inputDocPath, SaveFormat.Doc);

        // -----------------------------------------------------------------
        // Step 2: Load the DOC file that we just created.
        // -----------------------------------------------------------------
        Document docToConvert = new Document(inputDocPath);

        // -----------------------------------------------------------------
        // Step 3: Configure XlsxSaveOptions with default compression.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            // The default compression level is Normal; set explicitly for clarity.
            CompressionLevel = CompressionLevel.Normal,
            // XlsxSaveOptions must have SaveFormat set to Xlsx.
            SaveFormat = SaveFormat.Xlsx
        };

        // -----------------------------------------------------------------
        // Step 4: Convert and save the document as an XLSX workbook.
        // -----------------------------------------------------------------
        docToConvert.Save(outputXlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // Step 5: Validate that the output file was created successfully.
        // -----------------------------------------------------------------
        if (!File.Exists(outputXlsxPath))
        {
            throw new FileNotFoundException($"The output file '{outputXlsxPath}' was not created.");
        }

        FileInfo outputInfo = new FileInfo(outputXlsxPath);
        if (outputInfo.Length == 0)
        {
            throw new InvalidOperationException("The output XLSX file is empty.");
        }

        // Optional: Inform the user (no interactive input required).
        Console.WriteLine($"Conversion completed. Output file size: {outputInfo.Length} bytes.");
    }
}
