using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string basePath = Path.Combine(Directory.GetCurrentDirectory(), "BatchProcessingDemo");
        string inputFolder = Path.Combine(basePath, "InputDocs");
        string outputFolder = Path.Combine(basePath, "OutputDocs");

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample source documents (3 documents with a simple table each).
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Insert a heading.
            builder.Writeln($"Sample Document {i}");
            builder.Writeln();

            // Insert a 3x3 table.
            Table table = builder.StartTable();
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }
            builder.EndTable();

            // Save the sample document to the input folder.
            string inputPath = Path.Combine(inputFolder, $"Doc{i}.docx");
            sampleDoc.Save(inputPath);
        }

        // Process each document: replace all tables with a predefined style and margin settings.
        foreach (string inputPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(inputPath);

            // Iterate over all tables in the document.
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Apply a predefined style (built‑in "Table Grid").
                table.Style = doc.Styles["Table Grid"];

                // Set uniform cell margins (5 points on each side).
                foreach (Row row in table.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        cell.CellFormat.LeftPadding = 5;
                        cell.CellFormat.RightPadding = 5;
                        cell.CellFormat.TopPadding = 5;
                        cell.CellFormat.BottomPadding = 5;
                    }
                }
            }

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(inputPath));
            doc.Save(outputPath);
        }

        // Optional: verify that output files were created (no console output required).
        // The program ends here.
    }
}
