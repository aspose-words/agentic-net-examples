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

        // Define distinct table styles to apply.
        // Use StyleIdentifier enum – it contains the built‑in table style identifiers.
        StyleIdentifier[] styles = new StyleIdentifier[]
        {
            StyleIdentifier.TableGrid,
            StyleIdentifier.LightShadingAccent1,
            StyleIdentifier.LightListAccent2
        };

        // Build a table for each style.
        foreach (StyleIdentifier style in styles)
        {
            // Start a new table.
            builder.StartTable();

            // Create 2 rows with 3 columns each.
            for (int row = 1; row <= 2; row++)
            {
                for (int col = 1; col <= 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"Style: {style}, R{row}C{col}");
                }
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Retrieve the table that was just added.
            Table table = doc.LastSection.Body.Tables[doc.LastSection.Body.Tables.Count - 1];

            // Apply the distinct style.
            table.StyleIdentifier = style;

            // Add consistent spacing after the table.
            builder.Writeln(); // Inserts an empty paragraph as spacing.
        }

        // Save the document to a file.
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
