using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the intermediate DOCX and final XLSX files
        string docxPath = Path.Combine(artifactsDir, "Sample.docx");
        string xlsxPath = Path.Combine(artifactsDir, "Sample.xlsx");

        // Create a simple DOCX document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for XLSX conversion.");

        // Save the document as DOCX (bootstrap step)
        doc.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document
        Document loadedDoc = new Document(docxPath);

        // Configure XlsxSaveOptions with maximum compression
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions
        {
            CompressionLevel = CompressionLevel.Maximum,
            SaveFormat = SaveFormat.Xlsx // explicit, though default for XlsxSaveOptions
        };

        // Save the document as XLSX using the configured options
        loadedDoc.Save(xlsxPath, xlsxOptions);

        // Verify that the XLSX file was created
        if (!File.Exists(xlsxPath))
            throw new InvalidOperationException("The XLSX file was not created.");

        // Output file size (optional, for confirmation)
        long fileSize = new FileInfo(xlsxPath).Length;
        Console.WriteLine($"XLSX file created at '{xlsxPath}' with size {fileSize} bytes.");
    }
}
