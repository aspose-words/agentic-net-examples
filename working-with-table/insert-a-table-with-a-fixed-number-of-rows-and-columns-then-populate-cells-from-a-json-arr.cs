using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // JSON array representing rows and columns of the table.
        string json = @"[
            [""R1C1"", ""R1C2"", ""R1C3""],
            [""R2C1"", ""R2C2"", ""R2C3""],
            [""R3C1"", ""R3C2"", ""R3C3""]
        ]";

        // Deserialize JSON into a list of rows, each row is a list of cell strings.
        List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);

        if (tableData == null || tableData.Count == 0)
            throw new InvalidOperationException("JSON does not contain any table data.");

        // Determine the number of rows and columns.
        int rowCount = tableData.Count;
        int columnCount = tableData[0].Count;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the table.
        builder.StartTable();

        for (int i = 0; i < rowCount; i++)
        {
            // Ensure each row has the expected number of columns.
            if (tableData[i].Count != columnCount)
                throw new InvalidOperationException($"Row {i} does not have the expected number of columns.");

            for (int j = 0; j < columnCount; j++)
            {
                // Insert a new cell and write the corresponding text (no line break).
                builder.InsertCell();
                builder.Write(tableData[i][j]);
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table construction.
        builder.EndTable();

        // Define output file path.
        string outputPath = "OutputTable.docx";

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reopen the document and verify the first cell's text.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        Cell firstCell = loadedTable.Rows[0].Cells[0];
        string firstCellText = firstCell.GetText().TrimEnd('\a', '\r', '\n');

        if (firstCellText != tableData[0][0])
            throw new InvalidOperationException("Verification failed: first cell content does not match expected value.");

        // Execution completed successfully.
    }
}
