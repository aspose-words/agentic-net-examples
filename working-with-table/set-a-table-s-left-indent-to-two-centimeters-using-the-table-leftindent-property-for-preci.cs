using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Write("Sample cell");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Set the left indent of the table to 2 centimeters.
        // 1 centimeter = 28.3464567 points, so 2 cm ≈ 56.6929 points.
        table.LeftIndent = 2 * 28.3464567;

        // Save the document to the current directory.
        doc.Save("TableLeftIndent.docx");
    }
}
