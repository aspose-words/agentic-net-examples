using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with placeholder text.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputFolder, "Sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello _Name_! This is a sample document.");
        builder.Writeln("Date: _Date_");
        // Save the initial document.
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document and perform find‑and‑replace operations.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        // Replace placeholders with actual values.
        loadedDoc.Range.Replace("_Name_", "John Doe");
        loadedDoc.Range.Replace("_Date_", DateTime.Today.ToString("yyyy-MM-dd"));

        // -----------------------------------------------------------------
        // 3. Convert the modified document to Excel (XLSX) format.
        // -----------------------------------------------------------------
        string xlsxPath = Path.Combine(outputFolder, "Result.xlsx");
        // Simple conversion using the appropriate SaveFormat.
        loadedDoc.Save(xlsxPath, SaveFormat.Xlsx);

        // -----------------------------------------------------------------
        // 4. Validate that the XLSX file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(xlsxPath))
        {
            throw new InvalidOperationException($"The file '{xlsxPath}' was not created.");
        }

        Console.WriteLine("Find‑and‑replace completed and document saved as XLSX:");
        Console.WriteLine(xlsxPath);
    }
}
