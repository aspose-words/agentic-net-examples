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

        // First table with a caption.
        InsertTableWithCaption(builder, "First table caption", new string[,] {
            { "Header 1", "Header 2" },
            { "Row 1, Cell 1", "Row 1, Cell 2" },
            { "Row 2, Cell 1", "Row 2, Cell 2" }
        });

        // Second table with a caption.
        InsertTableWithCaption(builder, "Second table caption", new string[,] {
            { "Col A", "Col B", "Col C" },
            { "A1", "B1", "C1" },
            { "A2", "B2", "C2" },
            { "A3", "B3", "C3" }
        });

        // Update all fields (including the SEQ fields used for captions) so numbering is correct.
        doc.UpdateFields();

        // Define output file path.
        string outputPath = "TableWithCaptions.docx";

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }

    // Helper method to insert a caption and a table.
    private static void InsertTableWithCaption(DocumentBuilder builder, string captionText, string[,] data)
    {
        // Insert the caption: "Table {SEQ Table \* ARABIC}: <captionText>"
        builder.Write("Table ");
        builder.InsertField("SEQ Table \\* ARABIC");
        builder.Writeln($": {captionText}");

        // Start the table.
        Table table = builder.StartTable();

        // Build rows and cells from the provided 2D array.
        int rows = data.GetLength(0);
        int cols = data.GetLength(1);

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                builder.InsertCell();
                builder.Write(data[r, c]);
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Add a blank paragraph after the table for readability.
        builder.Writeln();
    }
}
