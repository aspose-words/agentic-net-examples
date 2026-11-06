using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample template document that contains a placeholder.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Dear {{Name}},");
        builder.Writeln("Thank you for your business.");
        templateDoc.Save("Template.docx");

        // ---------------------------------------------------------------
        // 2. Simulate retrieving data from a database.
        //    Here we use an in‑memory DataTable to avoid external dependencies.
        // ---------------------------------------------------------------
        string nameFromDb;
        DataTable users = new DataTable();
        users.Columns.Add("Id", typeof(int));
        users.Columns.Add("Name", typeof(string));
        users.Rows.Add(1, "Alice");

        // Retrieve the name where Id = 1.
        DataRow row = users.Select("Id = 1")[0];
        nameFromDb = row["Name"].ToString();

        // ---------------------------------------------------------------
        // 3. Load the template document.
        // ---------------------------------------------------------------
        Document doc = new Document("Template.docx");

        // ---------------------------------------------------------------
        // 4. Replace the placeholder with the value retrieved from the "database".
        // ---------------------------------------------------------------
        doc.Range.Replace("{{Name}}", nameFromDb, new FindReplaceOptions());

        // ---------------------------------------------------------------
        // 5. Save the resulting document.
        // ---------------------------------------------------------------
        doc.Save("Result.docx");
    }
}
