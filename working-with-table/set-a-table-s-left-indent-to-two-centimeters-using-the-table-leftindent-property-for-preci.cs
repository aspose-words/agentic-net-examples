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

        // Start a table and insert the first cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Set the left indent to 2 centimeters (1 cm = 28.3464567 points).
        double pointsPerCentimeter = 28.3464567;
        table.LeftIndent = 2 * pointsPerCentimeter;

        // Add some sample text to the cell.
        builder.Write("This table is indented by 2 cm from the left margin.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableLeftIndent.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
