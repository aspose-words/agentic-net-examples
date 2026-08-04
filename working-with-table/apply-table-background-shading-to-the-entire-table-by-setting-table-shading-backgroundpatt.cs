using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");

        // End the first row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Apply background shading to the entire table.
        // Use SetShading with a solid texture and the desired background color.
        table.SetShading(TextureIndex.TextureSolid, Color.Empty, Color.LightBlue);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableShading.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
