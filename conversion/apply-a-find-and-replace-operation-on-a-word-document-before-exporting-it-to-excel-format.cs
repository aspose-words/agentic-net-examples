using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document with placeholder text.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Hello <<Name>>, welcome to Aspose.Words!");
        source.Save("input.docx", SaveFormat.Docx);

        // Step 2: Load the created document.
        Document doc = new Document("input.docx");

        // Step 3: Perform find‑and‑replace operation.
        doc.Range.Replace("<<Name>>", "World", new FindReplaceOptions());

        // Step 4: Export the modified document to Excel format.
        doc.Save("output.xlsx", SaveFormat.Xlsx);

        // Step 5: Validate that the Excel file was created.
        if (!File.Exists("output.xlsx"))
        {
            throw new InvalidOperationException("Expected output Excel file was not created.");
        }
    }
}
