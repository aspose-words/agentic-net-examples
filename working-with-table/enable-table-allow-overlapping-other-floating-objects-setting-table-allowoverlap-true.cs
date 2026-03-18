using System;
using Aspose.Words;
using Aspose.Words.Tables;

class EnableTableOverlap
{
    static void Main()
    {
        // Load an existing Word document that contains at least one table.
        Document doc = new Document("Input.docx");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Ensure the table is a floating table (required for overlapping behavior).
        // Setting TextWrapping to Around makes the table float.
        table.TextWrapping = TextWrapping.Around;

        // Table.AllowOverlap is a read‑only property; its default value is true.
        // No setter is available, so we simply verify that overlapping is allowed.
        bool overlapAllowed = table.AllowOverlap; // should be true

        // (Optional) Output the current state for debugging purposes.
        Console.WriteLine($"Table AllowOverlap: {overlapAllowed}");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
