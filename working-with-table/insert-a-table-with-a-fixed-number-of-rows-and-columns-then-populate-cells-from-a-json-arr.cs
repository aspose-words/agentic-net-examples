using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeTableFromJson
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample JSON representing a rectangular table (rows of columns).
            // Each inner array is a row; each string is a cell value.
            string json = @"
            [
                [""Header 1"", ""Header 2"", ""Header 3""],
                [""Row1 Col1"", ""Row1 Col2"", ""Row1 Col3""],
                [""Row2 Col1"", ""Row2 Col2"", ""Row2 Col3""]
            ]";

            // Deserialize JSON into a list of rows, each row being a list of cell strings.
            List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);

            if (tableData == null || tableData.Count == 0)
                throw new InvalidOperationException("JSON does not contain any table data.");

            // Determine the maximum column count to ensure a uniform table.
            int columnCount = 0;
            foreach (var row in tableData)
                if (row != null && row.Count > columnCount)
                    columnCount = row.Count;

            // Begin building the table.
            builder.StartTable();

            foreach (var row in tableData)
            {
                // Ensure each row has the same number of cells.
                for (int col = 0; col < columnCount; col++)
                {
                    builder.InsertCell();

                    string cellText = (row != null && col < row.Count) ? row[col] : string.Empty;
                    builder.Write(cellText);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to a file in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
