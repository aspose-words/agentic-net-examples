using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Access the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Get the total number of rows in the table.
        int rowCount = table.Rows.Count;

        // Get the total number of columns by counting cells in the first row.
        int columnCount = table.Rows[0].Cells.Count;

        // Display the row and column counts.
        Console.WriteLine($"Rows: {rowCount}, Columns: {columnCount}");

        // Save the (unchanged) document to a new file.
        doc.Save("Output.docx");
    }
}
