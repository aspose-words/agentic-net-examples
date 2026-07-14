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

        // Start a simple table with one cell.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Set the left indent of the table to 2 centimeters.
        // 1 inch = 72 points, 1 cm = 72 / 2.54 points.
        double pointsPerCentimeter = 72.0 / 2.54;
        table.LeftIndent = 2 * pointsPerCentimeter; // ≈ 56.7 points

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableLeftIndent.docx");
        doc.Save(outputPath);
    }
}
