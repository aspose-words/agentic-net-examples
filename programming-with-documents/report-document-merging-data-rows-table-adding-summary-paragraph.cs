using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class ReportGenerator
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title for the report.
        builder.Writeln("Employee Report");
        builder.Writeln();

        // Build a table with a mail‑merge region named "Employees".
        // The region will be repeated for each row in the data source.
        builder.StartTable();

        // First cell – start of the region.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableStart:Employees");

        // Header row (optional – static text, not part of the region).
        builder.InsertCell();
        builder.Write("Name");
        builder.InsertCell();
        builder.Write("Age");
        builder.InsertCell();
        builder.Write("Country");
        builder.EndRow();

        // Data row – contains merge fields that will be populated.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Name");
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Age");
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Country");
        // End of the region.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableEnd:Employees");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln(); // Add a blank line after the table.

        // Create a DataTable that matches the merge fields.
        DataTable employeeTable = new DataTable("Employees");
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Age", typeof(int));
        employeeTable.Columns.Add("Country", typeof(string));

        // Populate the DataTable with sample data.
        employeeTable.Rows.Add("John Doe", 30, "USA");
        employeeTable.Rows.Add("Anna Smith", 27, "UK");
        employeeTable.Rows.Add("Li Wei", 35, "China");

        // Execute the mail merge with regions – this will repeat the data row for each record.
        doc.MailMerge.ExecuteWithRegions(employeeTable);

        // Add a summary paragraph after the table.
        builder.MoveToDocumentEnd();
        builder.Writeln($"Summary: Total employees = {employeeTable.Rows.Count}");

        // Save the generated report.
        doc.Save("EmployeeReport.docx");
    }
}
