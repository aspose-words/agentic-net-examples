using System;
using System.Data;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that contains a placeholder to be replaced.
        builder.Writeln("Hello _FullName_, welcome to our company.");

        // Simulate retrieving a value from a database.
        DataTable dbTable = new DataTable();
        dbTable.Columns.Add("FullName", typeof(string));
        dbTable.Rows.Add("John Doe"); // Example data row.

        // Extract the dynamic value.
        string fullName = dbTable.Rows[0]["FullName"].ToString();

        // Replace the placeholder text in the whole document range.
        doc.Range.Replace("_FullName_", fullName);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
