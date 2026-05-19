using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static footer that will appear on every page.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Confidential – Do not distribute");
        // Return the builder to the main story to continue building the body.
        builder.MoveToDocumentEnd();

        // Insert mail‑merge fields into the document body.
        builder.Writeln();
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare data for the mail merge.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello! This is a mail merge example." };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        doc.Save("MailMergeWithFooter.docx");
    }
}
