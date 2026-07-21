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

        // Start a table and add a single cell with some text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell content.");
        builder.EndTable();

        // Set the left indent of the table to 1 cm (≈28.35 points).
        table.LeftIndent = 28.35; // 1 cm in points.

        // Table.RightIndent does not exist in this API version.
        // Use DistanceRight to achieve a similar effect: the space between the table and surrounding text.
        table.DistanceRight = 28.35; // Approximate right margin of 1 cm.

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableMargins.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
