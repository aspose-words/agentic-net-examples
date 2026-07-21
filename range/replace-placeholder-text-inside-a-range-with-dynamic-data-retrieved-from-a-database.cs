using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // ---------------------------------------------------------------
        // 1. Simulate a data source (e.g., a database) using a DataTable.
        // ---------------------------------------------------------------
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("Id", typeof(int));
        employees.Columns.Add("FullName", typeof(string));

        // Insert a sample record.
        employees.Rows.Add(1, "John Doe");

        // Retrieve the FullName value for Id = 1.
        string fullName = null;
        foreach (DataRow row in employees.Rows)
        {
            if ((int)row["Id"] == 1)
            {
                fullName = row["FullName"].ToString();
                break;
            }
        }

        // ---------------------------------------------------------------
        // 2. Create a Word document containing a placeholder to be replaced.
        // ---------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Employee Report");
        builder.Writeln("Name: {{FullName}}"); // Placeholder that will be replaced.

        // ---------------------------------------------------------------
        // 3. Replace the placeholder text inside the document's range.
        // ---------------------------------------------------------------
        // Use the Range.Replace method for a simple string replacement.
        // The placeholder "{{FullName}}" is replaced with the value retrieved from the data source.
        doc.Range.Replace("{{FullName}}", fullName ?? string.Empty);

        // ---------------------------------------------------------------
        // 4. Save the resulting document to the local file system.
        // ---------------------------------------------------------------
        const string outputPath = "EmployeeReport.docx";
        doc.Save(outputPath);
    }
}
