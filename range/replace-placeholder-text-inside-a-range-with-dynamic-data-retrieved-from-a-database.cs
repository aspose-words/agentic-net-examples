using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Simulate a database by using an in‑memory DataTable.
        DataTable table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));

        // Define the primary key so that Find can locate rows by Id.
        table.PrimaryKey = new[] { table.Columns["Id"] };

        // Insert a sample record.
        table.Rows.Add(1, "John Doe");

        // Retrieve the name for replacement (mimicking a DB query).
        string nameFromDb = table.Rows.Find(1) != null
            ? table.Rows.Find(1)["Name"].ToString()
            : "Unknown";

        // Build a Word document containing a placeholder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dear _FullName_,");
        builder.Writeln("Welcome to our service.");

        // Replace the placeholder with the value retrieved from the simulated database.
        doc.Range.Replace("_FullName_", nameFromDb, new FindReplaceOptions());

        // Save the resulting document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Output the final text to the console for verification.
        Console.WriteLine("Document text after replacement:");
        Console.WriteLine(doc.GetText().Trim());
    }
}
