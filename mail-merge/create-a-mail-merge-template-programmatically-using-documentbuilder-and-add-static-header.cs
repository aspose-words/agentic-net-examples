using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeTemplateExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a static header that will appear on every page.
            // Using a header section ensures the text is repeated automatically.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 16;
            builder.Font.Bold = true;
            builder.Writeln("Customer Invoice");

            // Return to the main document body.
            builder.MoveToDocumentEnd();

            // Insert a line break after the header.
            builder.Writeln();

            // Insert merge fields that will be filled during a mail merge.
            builder.Font.Size = 12;
            builder.Font.Bold = false;
            builder.Writeln("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(",");

            builder.Writeln();
            builder.Writeln("Thank you for your purchase. Below are the details of your order:");
            builder.Writeln();

            // Example table with merge fields for order items.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.InsertCell();
            builder.Write("Price");
            builder.EndRow();

            // First row (placeholder data – actual data will be supplied by mail merge).
            builder.InsertCell();
            builder.InsertField("MERGEFIELD ItemName", "<ItemName>");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD Quantity", "<Quantity>");
            builder.InsertCell();
            builder.InsertField("MERGEFIELD Price", "<Price>");
            builder.EndRow();

            builder.EndTable();

            builder.Writeln();
            builder.Writeln("Total: ");
            builder.InsertField("MERGEFIELD TotalAmount", "<TotalAmount>");

            // Save the template to disk.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "MailMergeTemplate.docx");
            doc.Save(outputPath);
        }
    }
}
