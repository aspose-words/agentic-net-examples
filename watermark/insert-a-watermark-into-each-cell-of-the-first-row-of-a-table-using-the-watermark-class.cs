using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a first row with three cells.
        builder.StartTable();

        // Insert three cells in the first row.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Cell {i + 1}");
        }

        // End the first row.
        builder.EndRow();

        // Add a second row to demonstrate normal table content.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Data {i + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a text watermark to the whole document.
        // The watermark will appear behind the text in every cell, including each cell of the first row.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Save the document to a local file.
        const string outputPath = "WatermarkedTable.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
