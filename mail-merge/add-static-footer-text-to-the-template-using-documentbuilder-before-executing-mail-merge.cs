using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Add a static footer that will appear on every page.
        // ------------------------------------------------------------
        // Move the builder to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        // Write the static text.
        builder.Write("Confidential – For internal use only");
        // Add a line break after the footer text.
        builder.Writeln();

        // ------------------------------------------------------------
        // Build a simple mail‑merge template.
        // ------------------------------------------------------------
        // Move back to the main body of the document.
        builder.MoveToDocumentEnd();
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // ------------------------------------------------------------
        // Prepare sample data for the mail merge.
        // ------------------------------------------------------------
        DataTable table = new DataTable("Data");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Message");

        table.Rows.Add("John", "Doe", "Welcome to Aspose.Words!");
        table.Rows.Add("Jane", "Smith", "Your order has been shipped.");

        // ------------------------------------------------------------
        // Execute the mail merge.
        // ------------------------------------------------------------
        doc.MailMerge.Execute(table);

        // ------------------------------------------------------------
        // Save the resulting document.
        // ------------------------------------------------------------
        doc.Save("MailMergeWithFooter.docx");
    }
}
