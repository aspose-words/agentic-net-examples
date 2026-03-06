using System;
using System.IO;
using Aspose.Words;

class MailMergeTemplateExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple mail‑merge template with MERGEFIELD fields.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeTemplate.docx");

        // Save the template as a DOCX file.
        doc.Save(outputPath);
    }
}
