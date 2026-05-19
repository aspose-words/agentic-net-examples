using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static header centered and bold.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Customer Invoice");

        // Add an empty line after the header.
        builder.Writeln();

        // Insert merge fields for the mail‑merge template.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Save the template to a file.
        doc.Save("MailMergeTemplate.docx");
    }
}
