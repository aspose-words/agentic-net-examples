using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace AsposeWordsTableFromJson
{
    public class Program
    {
        public static void Main()
        {
            // JSON array representing rows and columns of the table.
            // Each inner array is a row, and each string is a cell value.
            string json = @"
            [
                [""R1C1"", ""R1C2"", ""R1C3""],
                [""R2C1"", ""R2C2"", ""R2C3""],
                [""R3C1"", ""R3C2"", ""R3C3""]
            ]";

            // Deserialize the JSON into a list of rows, each row being a list of cell strings.
            List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(json);

            // Validate that we have at least one row and one column.
            if (tableData == null || tableData.Count == 0 || tableData[0].Count == 0)
                throw new InvalidOperationException("JSON does not contain a valid table structure.");

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            Table table = builder.StartTable();

            // Iterate over each row.
            foreach (List<string> row in tableData)
            {
                // Iterate over each cell in the current row.
                foreach (string cellText in row)
                {
                    // Insert a new cell and write its content.
                    builder.InsertCell();
                    builder.Write(cellText);
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            const string outputFileName = "TableFromJson.docx";
            doc.Save(outputFileName);
        }
    }
}
