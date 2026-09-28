using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // JSON representing a 2‑dimensional array (rows and columns).
        string json = @"[
            [""Header1"", ""Header2"", ""Header3""],
            [""Row1Col1"", ""Row1Col2"", ""Row1Col3""],
            [""Row2Col1"", ""Row2Col2"", ""Row2Col3""]
        ]";

        // Deserialize the JSON into a list of rows, each row being a list of cell strings.
        List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);
        if (tableData == null || tableData.Count == 0)
            throw new Exception("The JSON does not contain any table data.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building the table.
        builder.StartTable();

        // Populate the table cells from the deserialized JSON data.
        foreach (List<string> row in tableData)
        {
            foreach (string cellText in row)
            {
                builder.InsertCell();
                builder.Writeln(cellText);
            }
            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "OutputTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
