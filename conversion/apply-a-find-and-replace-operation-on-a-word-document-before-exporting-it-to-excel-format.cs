using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with placeholder text.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Dear _Customer_,");
        builder.Writeln("Thank you for your purchase of _Product_.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the created document.
        Document doc = new Document("input.docx");

        // Perform find‑and‑replace operations.
        doc.Range.Replace("_Customer_", "John Doe");
        doc.Range.Replace("_Product_", "Aspose.Words Library");

        // Export the modified document to Excel format.
        doc.Save("output.xlsx", SaveFormat.Xlsx);

        // Verify that the Excel file was created.
        if (!File.Exists("output.xlsx"))
            throw new InvalidOperationException("The expected Excel output file was not created.");
    }
}
