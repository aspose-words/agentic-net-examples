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

        // Return the builder to the main body of the document.
        builder.MoveToDocumentEnd();

        // Insert simple mail‑merge fields into the body.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(".");

        // Data to be merged.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Perform the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        doc.Save("Output.docx");
    }
}
