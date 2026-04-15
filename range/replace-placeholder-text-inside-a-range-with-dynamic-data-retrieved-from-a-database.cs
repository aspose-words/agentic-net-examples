using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace AsposeWordsRangeReplaceDemo
{
    public class Program
    {
        public static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Insert a paragraph that contains a placeholder text.
            //    The placeholder will be replaced later with data from the "database".
            const string placeholder = "_FullName_";
            builder.Writeln($"Hello, {placeholder}! Welcome to our service.");

            // 3. Simulate a database by creating an in‑memory DataTable.
            //    In a real scenario this could be replaced with an actual DB query.
            DataTable dbTable = new DataTable("Customers");
            dbTable.Columns.Add("FullName", typeof(string));
            dbTable.Rows.Add("Jane Doe"); // Sample data row.

            // 4. Retrieve the dynamic value from the simulated database.
            string fullName = dbTable.Rows[0]["FullName"].ToString();

            // 5. Perform a find‑and‑replace operation on the whole‑document range.
            //    The Replace method is case‑insensitive by default.
            int replacements = doc.Range.Replace(placeholder, fullName);

            // Optional: verify that the replacement occurred.
            Console.WriteLine($"Replacements made: {replacements}");
            Console.WriteLine("Resulting document text:");
            Console.WriteLine(doc.GetText().Trim());

            // 6. Save the modified document to the local file system.
            const string outputPath = "OutputDocument.docx";
            doc.Save(outputPath);
        }
    }
}
