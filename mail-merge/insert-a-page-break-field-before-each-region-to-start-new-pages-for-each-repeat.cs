using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a page break before the mail merge region so each region starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Define a mail merge region named "Employees".
        builder.InsertField(" MERGEFIELD TableStart:Employees");

        // Add some content inside the region.
        builder.Write("Employee Name: ");
        builder.InsertField(" MERGEFIELD Name");
        builder.Writeln();

        // End the mail merge region.
        builder.InsertField(" MERGEFIELD TableEnd:Employees");

        // Prepare a DataTable that will be used for the mail merge.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Name");
        employees.Rows.Add("John Doe");
        employees.Rows.Add("Jane Smith");
        employees.Rows.Add("Bob Johnson");

        // Perform mail merge with regions. The region will be repeated for each row,
        // and because we placed a page break before the region, each repeat starts on a new page.
        doc.MailMerge.ExecuteWithRegions(employees);

        // Save the resulting document.
        doc.Save("MailMergeWithPageBreak.docx");
    }
}
