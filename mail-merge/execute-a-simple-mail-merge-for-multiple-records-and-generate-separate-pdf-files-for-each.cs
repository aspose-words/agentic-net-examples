using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a mail‑merge template document in memory.
        Document template = CreateTemplate();

        // Build a data source with several records.
        DataTable data = BuildDataTable();

        // Perform a separate mail merge for each record and save each result as a PDF.
        for (int i = 0; i < data.Rows.Count; i++)
        {
            // Clone the template so each record starts from the original layout.
            Document mergedDoc = (Document)template.Clone();

            // Merge the current row into the document.
            mergedDoc.MailMerge.Execute(data.Rows[i]);

            // Save the merged document as a PDF file.
            string fileName = $"MergedDocument_{i + 1}.pdf";
            mergedDoc.Save(fileName, SaveFormat.Pdf);
        }
    }

    // Creates a simple document containing three MERGEFIELDs.
    private static Document CreateTemplate()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        return doc;
    }

    // Constructs a DataTable with sample data for the mail merge.
    private static DataTable BuildDataTable()
    {
        DataTable table = new DataTable("Recipients");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Message");

        table.Rows.Add("John", "Doe", "Hello! This is the first message.");
        table.Rows.Add("Jane", "Smith", "Greetings from the second record.");
        table.Rows.Add("Bob", "Johnson", "This is the third example message.");

        return table;
    }
}
