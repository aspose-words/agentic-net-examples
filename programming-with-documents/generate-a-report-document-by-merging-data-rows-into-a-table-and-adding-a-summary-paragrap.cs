using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class ReportGenerator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge region start field.
        builder.InsertField(" MERGEFIELD TableStart:Data ");

        // Build a simple two‑column table with merge fields for each column.
        builder.StartTable();

        // Header row (optional – not part of the merge region).
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data row – contains the fields that will be populated by the merge.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Product ");
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Quantity ");
        builder.EndRow();

        builder.EndTable();

        // Insert the mail‑merge region end field.
        builder.InsertField(" MERGEFIELD TableEnd:Data ");

        // Prepare a DataTable that will serve as the data source for the merge.
        DataTable data = new DataTable("Data");
        data.Columns.Add("Product", typeof(string));
        data.Columns.Add("Quantity", typeof(int));

        data.Rows.Add("Apples", 120);
        data.Rows.Add("Bananas", 85);
        data.Rows.Add("Cherries", 60);

        // Execute the mail merge with regions – this will repeat the table row for each DataRow.
        doc.MailMerge.ExecuteWithRegions(data);

        // After the merge, add a summary paragraph that shows the total number of rows.
        DocumentBuilder summaryBuilder = new DocumentBuilder(doc);
        summaryBuilder.MoveToDocumentEnd();
        summaryBuilder.Writeln();
        summaryBuilder.Writeln($"Total products listed: {data.Rows.Count}");

        // Save the final report to the local file system.
        doc.Save("Report.docx");
    }
}
