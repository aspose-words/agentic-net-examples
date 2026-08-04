using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add static header text.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Company Confidential - Mail Merge Template");
        builder.Writeln(); // Move to next line in the header.

        // Return to the main body of the document.
        builder.MoveToDocumentEnd();

        // Build the mail‑merge body with static text and MERGEFIELDs.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Save the template to a file.
        const string outputPath = "MailMergeTemplate.docx";
        doc.Save(outputPath);
    }
}
