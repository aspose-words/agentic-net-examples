using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some placeholder text that will be replaced later.
        builder.Writeln("Dear _FullName_,");
        builder.Writeln("Welcome to _Company_.");
        builder.Writeln("Your employee ID is _EmployeeID_.");

        // Simulate retrieving data from a database.
        // In a real scenario this DataTable would be filled by a DB query.
        DataTable data = new DataTable();
        data.Columns.Add("FullName", typeof(string));
        data.Columns.Add("Company", typeof(string));
        data.Columns.Add("EmployeeID", typeof(string));
        data.Rows.Add("John Doe", "Acme Corp", "A12345");

        // Extract the values from the first (and only) row.
        DataRow row = data.Rows[0];
        string fullName = row["FullName"].ToString();
        string company = row["Company"].ToString();
        string employeeId = row["EmployeeID"].ToString();

        // Replace the placeholders in the whole document range.
        // The Replace method performs a case‑insensitive search.
        doc.Range.Replace("_FullName_", fullName);
        doc.Range.Replace("_Company_", company);
        doc.Range.Replace("_EmployeeID_", employeeId);

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
