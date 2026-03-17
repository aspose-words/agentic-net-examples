using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Set a preferred width so the table can float.
        table.PreferredWidth = PreferredWidth.FromPoints(300);

        // Configure the table to wrap text around it.
        table.TextWrapping = TextWrapping.Around;

        // Optional: define distances between the table and surrounding text.
        table.DistanceLeft = 24;
        table.DistanceRight = 24;
        table.DistanceTop = 3;
        table.DistanceBottom = 3;

        // Add surrounding text to demonstrate the wrapping effect.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document.
        doc.Save("TableWrapAround.docx");
    }
}
