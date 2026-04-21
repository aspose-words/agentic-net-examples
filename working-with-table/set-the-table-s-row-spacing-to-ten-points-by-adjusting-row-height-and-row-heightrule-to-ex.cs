using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add the first row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Row 1, Cell 1.");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2.");
        builder.EndRow();

        // Set the height of the next row to exactly 10 points.
        builder.RowFormat.Height = 10.0;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Add the second row which will inherit the height settings above.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1.");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RowSpacing.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
