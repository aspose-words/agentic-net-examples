using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a placeholder that will be replaced.
        builder.Writeln("Hello {{Name}}! Welcome to Aspose.Words.");

        // ---------- Simulate database retrieval ----------
        // Use an in‑memory DataTable to mimic a simple database table.
        string retrievedName;
        DataTable usersTable = new DataTable();
        usersTable.Columns.Add("Id", typeof(int));
        usersTable.Columns.Add("Name", typeof(string));

        // Insert a sample record.
        usersTable.Rows.Add(1, "John Doe");

        // Retrieve the name value where Id = 1.
        DataRow[] foundRows = usersTable.Select("Id = 1");
        if (foundRows.Length > 0)
        {
            retrievedName = foundRows[0]["Name"].ToString();
        }
        else
        {
            retrievedName = string.Empty;
        }

        // ---------- Replace placeholder text using Range ----------
        // Use Document.Range.Replace for a simple find‑and‑replace operation.
        FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
        doc.Range.Replace("{{Name}}", retrievedName, options);

        // Save the modified document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Optional: output the final text to the console for verification.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
        Console.WriteLine("Final document text:");
        Console.WriteLine(doc.Range.Text);
    }
}
