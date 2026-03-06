using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class MailMergeTemplateGenerator
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add merge fields to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a greeting line with MERGEFIELD placeholders.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Save the template as a DOCX file.
        doc.Save("MailMergeTemplate.docx");
    }
}
