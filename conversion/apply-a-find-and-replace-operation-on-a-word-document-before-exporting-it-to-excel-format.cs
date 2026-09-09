using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "input.docx";
        const string outputPath = "output.xlsx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample Word document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Thank you for your purchase.");
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Step 2: Load the document, perform find‑and‑replace.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);
        int replacements = doc.Range.Replace("_Customer_", "John Doe");
        if (replacements == 0)
            throw new InvalidOperationException("No occurrences were replaced.");

        // -----------------------------------------------------------------
        // Step 3: Export the modified document to Excel format.
        // -----------------------------------------------------------------
        doc.Save(outputPath, SaveFormat.Xlsx);

        // -----------------------------------------------------------------
        // Step 4: Validate that the Excel file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected Excel file was not created.");

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(inputPath);
        // File.Delete(outputPath);
    }
}
