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

        // Start a table.
        builder.StartTable();

        // Build a 2x2 table and set cell padding to 5 points on all sides.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");

                // The current paragraph resides inside the newly created cell.
                Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
                cell.CellFormat.LeftPadding = 5.0;
                cell.CellFormat.RightPadding = 5.0;
                cell.CellFormat.TopPadding = 5.0;
                cell.CellFormat.BottomPadding = 5.0;
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithPadding.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Indicate success (no interactive input required).
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
