using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static header centered, bold and larger font.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Customer Invoice");

        // Add an empty line after the header.
        builder.Writeln();

        // Insert merge fields for a simple mail‑merge template.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

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
