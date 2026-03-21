using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main(string[] args)
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample text");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Apply a uniform border thickness of 2 points to all sides of the table.
        table.SetBorders(LineStyle.Single, 2.0, Color.Black);

        // Save the modified document.
        doc.Save("Output.docx");
        Console.WriteLine("Document saved as Output.docx");
    }
}
