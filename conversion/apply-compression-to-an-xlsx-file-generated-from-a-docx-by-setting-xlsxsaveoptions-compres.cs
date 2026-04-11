using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the intermediate DOCX and final XLSX files.
        string docxPath = Path.Combine(artifactsDir, "Sample.docx");
        string xlsxPath = Path.Combine(artifactsDir, "Compressed.xlsx");

        // -----------------------------------------------------------------
        // Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for XLSX conversion.");
        // Save the DOCX locally as required by the input bootstrap rules.
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the DOCX file that we just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // Configure XlsxSaveOptions with Fast compression.
        // -----------------------------------------------------------------
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Fast,
            SaveFormat = SaveFormat.Xlsx
        };

        // Save the document as XLSX using the specified options.
        loadedDoc.Save(xlsxPath, xlsxOptions);

        // -----------------------------------------------------------------
        // Validate that the XLSX file exists and contains data.
        // -----------------------------------------------------------------
        FileInfo fileInfo = new FileInfo(xlsxPath);
        if (!fileInfo.Exists || fileInfo.Length == 0)
        {
            throw new InvalidOperationException("Failed to create a compressed XLSX file.");
        }

        // Output the result (non‑interactive).
        Console.WriteLine($"Compressed XLSX file created at: {xlsxPath}");
        Console.WriteLine($"File size: {fileInfo.Length} bytes");
    }
}
