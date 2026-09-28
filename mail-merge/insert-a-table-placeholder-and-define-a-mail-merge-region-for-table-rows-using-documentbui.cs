using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple greeting with merge fields.
        builder.Writeln("Dear <<Title>> <<LastName>>,");
        builder.Writeln();

        // Insert the start of the mail merge region for a table.
        builder.InsertField("TableStart:Employees", null);
        builder.Writeln();

        // Begin a table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.EndRow();

        // Data row with merge fields.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Name \\* MERGEFORMAT", null);
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Age \\* MERGEFORMAT", null);
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Insert the end of the mail merge region.
        builder.InsertField("TableEnd:Employees", null);
        builder.Writeln();

        // Prepare data for the mail merge region.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Columns.Add("Age");
        table.Rows.Add("John Doe", "30");
        table.Rows.Add("Jane Smith", "25");
        table.Rows.Add("Bob Johnson", "40");

        // Execute mail merge with the defined region.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
