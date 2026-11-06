using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Hello World! This is a sample document.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the created document.
        Document doc = new Document("input.docx");

        // Perform a find‑and‑replace: replace "Hello" with "Hi".
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc.Range.Replace("Hello", "Hi", replaceOptions);

        // Export the modified document to Excel format (XLSX).
        doc.Save("output.xlsx", SaveFormat.Xlsx);

        // Validate that the Excel file was created.
        if (!File.Exists("output.xlsx"))
        {
            throw new InvalidOperationException("Expected output XLSX was not created.");
        }

        // Optional: clean up temporary files (not required, but keeps the folder tidy).
        // File.Delete("input.docx");
    }
}
