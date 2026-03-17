using System;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Tables;

class TableFromJson
{
    static void Main()
    {
        // Load JSON data from a file. The JSON should be an array of objects, e.g.:
        // [
        //   { "Name": "Alice", "Age": 30, "City": "London" },
        //   { "Name": "Bob",   "Age": 25, "City": "Paris" }
        // ]
        string jsonPath = "data.json";
        string jsonContent = File.ReadAllText(jsonPath);
        JsonDocument jsonDoc = JsonDocument.Parse(jsonContent);
        JsonElement root = jsonDoc.RootElement;

        if (root.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("Root JSON element must be an array.");

        // Determine column names from the first object in the array.
        JsonElement firstItem = root[0];
        var columnNames = new string[firstItem.EnumerateObject().Count()];
        int colIndex = 0;
        foreach (var prop in firstItem.EnumerateObject())
            columnNames[colIndex++] = prop.Name;

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // ---- Header row ----
        foreach (string colName in columnNames)
        {
            builder.InsertCell();
            builder.Write(colName);
        }
        builder.EndRow();

        // ---- Data rows ----
        foreach (JsonElement item in root.EnumerateArray())
        {
            foreach (string colName in columnNames)
            {
                builder.InsertCell();
                // Retrieve the value for the current column; if missing, write empty string.
                if (item.TryGetProperty(colName, out JsonElement value))
                {
                    // Convert the JSON value to a string representation.
                    string text = value.ValueKind switch
                    {
                        JsonValueKind.String => value.GetString(),
                        JsonValueKind.Number => value.GetRawText(),
                        JsonValueKind.True => "True",
                        JsonValueKind.False => "False",
                        JsonValueKind.Null => "",
                        _ => value.GetRawText()
                    };
                    builder.Write(text);
                }
                else
                {
                    builder.Write(string.Empty);
                }
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("TableFromJson.docx");
    }
}
