using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeTemplateCreator
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "MailMergeTemplate.docx");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static header to the document.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Acme Corporation");
        builder.Writeln();
        builder.Write("123 Business Rd., Metropolis");
        builder.Writeln();
        builder.Write("Phone: (555) 123‑4567");
        builder.Writeln();
        builder.Font.Size = 14;
        builder.Font.Bold = true;
        builder.Writeln("Customer Invoice");
        builder.Font.Size = 12;
        builder.Font.Bold = false;

        // Return to the main body of the document.
        builder.MoveToDocumentEnd();

        // Insert some static text.
        builder.Writeln("Dear ");
        // Insert a MERGEFIELD for the customer's name.
        builder.InsertField("MERGEFIELD CustomerName", "<CustomerName>");
        builder.Writeln(",");

        builder.Writeln();
        builder.Writeln("Thank you for your purchase. Below are the details of your order:");

        // Insert a table with merge fields for product information.
        builder.StartTable();

        // Header row (static text).
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data row (merge fields).
        builder.InsertCell();
        builder.InsertField("MERGEFIELD ProductName", "<ProductName>");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity", "<Quantity>");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Price", "<Price>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln();
        builder.Write("Total: ");
        builder.InsertField("MERGEFIELD TotalPrice", "<TotalPrice>");
        builder.Writeln();

        builder.Writeln("Sincerely,");
        builder.Writeln("Acme Sales Team");

        // Save the template document.
        doc.Save(outputPath);
    }
}
