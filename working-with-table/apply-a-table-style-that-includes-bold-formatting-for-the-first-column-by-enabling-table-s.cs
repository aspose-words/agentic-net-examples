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

        // Build a 3x3 table.
        builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in style.
        table.Style = doc.Styles["Table Grid"];

        // Enable first‑column formatting in the style options.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Make the text in the first column bold.
        foreach (Row row in table.Rows)
        {
            Cell firstCell = row.Cells[0];
            Paragraph para = firstCell.FirstParagraph;
            if (para != null)
            {
                if (para.Runs.Count > 0)
                {
                    foreach (Run run in para.Runs)
                        run.Font.Bold = true;
                }
                else
                {
                    Run run = new Run(doc);
                    run.Font.Bold = true;
                    para.AppendChild(run);
                }
            }
        }

        // Save the document.
        string outputPath = "TableStyleFirstColumnBold.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
