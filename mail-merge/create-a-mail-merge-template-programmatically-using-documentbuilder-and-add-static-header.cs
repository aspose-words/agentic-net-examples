using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static header that will appear on every page.
        // Use the built‑in header/footer feature.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Customer Invoice");

        // Return to the main body of the document.
        builder.MoveToDocumentEnd();

        // Insert a blank line after the header.
        builder.Writeln();

        // Add merge fields that will be filled during a mail merge.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.Writeln();

        builder.Write("Thank you for your purchase of ");
        builder.InsertField("MERGEFIELD ProductName", "<ProductName>");
        builder.Write(" on ");
        builder.InsertField("MERGEFIELD PurchaseDate", "<PurchaseDate>");
        builder.Writeln(".");

        // Save the template to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeTemplate.docx");
        doc.Save(outputPath);
    }
}
