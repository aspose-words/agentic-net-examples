using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MailMergeCleaner
{
    public static void Execute()
    {
        // Load the template document that contains MERGEFIELDs or mail‑merge regions.
        Document doc = new Document("Template.docx");

        // Configure mail‑merge to remove any empty paragraphs that may be created.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

        // Optional: treat paragraphs that contain only punctuation as empty as well.
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // ---------- First data source ----------
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("FirstName");
        customers.Columns.Add("LastName");
        customers.Rows.Add("John", "Doe");
        customers.Rows.Add("", ""); // This row would generate an empty paragraph.

        // Execute mail merge for the first region (TableStart:Customers / TableEnd:Customers).
        doc.MailMerge.ExecuteWithRegions(customers);

        // ---------- Second data source ----------
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderId");
        orders.Columns.Add("Product");
        orders.Rows.Add("001", "Widget");
        orders.Rows.Add("", ""); // Empty row → empty paragraph to be removed.

        // Execute mail merge for the second region (TableStart:Orders / TableEnd:Orders).
        doc.MailMerge.ExecuteWithRegions(orders);

        // Save the final document; empty paragraphs have been removed.
        doc.Save("CleanedResult.docx");
    }

    // Entry point required by the C# compiler.
    static void Main(string[] args)
    {
        Execute();
    }
}
