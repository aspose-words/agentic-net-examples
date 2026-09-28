using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ReportWithMultipleTables
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the built‑in table styles to be applied.
        StyleIdentifier[] styles = new StyleIdentifier[]
        {
            StyleIdentifier.TableGrid,
            StyleIdentifier.LightShadingAccent1,
            StyleIdentifier.MediumShading1Accent2
        };

        // Build a table for each style.
        foreach (StyleIdentifier style in styles)
        {
            // Start a new table.
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Retrieve the table that was just created.
            Table table = doc.LastSection.Body.Tables[doc.LastSection.Body.Tables.Count - 1];

            // Apply the distinct style.
            table.StyleIdentifier = style;

            // Add an empty paragraph to create consistent spacing between tables.
            builder.Writeln();
        }

        // Define output file path.
        string outputPath = "ReportWithTables.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine("Document created successfully.");
    }
}
