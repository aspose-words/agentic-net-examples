using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Tables;

class TableFromJson
{
    static void Main()
    {
        string jsonPath = "data.json";
        string jsonContent;

        if (File.Exists(jsonPath))
        {
            jsonContent = File.ReadAllText(jsonPath);
        }
        else
        {
            // Fallback sample JSON data (array of objects)
            jsonContent = @"[
                { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""London"" },
                { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""Paris"" }
            ]";
        }

        using JsonDocument jsonDoc = JsonDocument.Parse(jsonContent);
        JsonElement root = jsonDoc.RootElement;

        // Determine the collection of items (array or single object)
        IEnumerable<JsonElement> items;
        if (root.ValueKind == JsonValueKind.Array)
        {
            items = root.EnumerateArray();
        }
        else if (root.ValueKind == JsonValueKind.Object)
        {
            items = new[] { root };
        }
        else
        {
            throw new InvalidOperationException("Root JSON element must be an array or an object.");
        }

        // Get the first item to determine column names
        using IEnumerator<JsonElement> enumerator = items.GetEnumerator();
        if (!enumerator.MoveNext())
            throw new InvalidOperationException("JSON data contains no items.");

        JsonElement firstItem = enumerator.Current;
        var columnNames = new List<string>();
        foreach (var prop in firstItem.EnumerateObject())
            columnNames.Add(prop.Name);

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
        // Write the first item
        WriteRow(builder, firstItem, columnNames);

        // Write remaining items
        while (enumerator.MoveNext())
        {
            WriteRow(builder, enumerator.Current, columnNames);
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("TableFromJson.docx");
    }

    private static void WriteRow(DocumentBuilder builder, JsonElement item, List<string> columnNames)
    {
        foreach (string colName in columnNames)
        {
            builder.InsertCell();
            if (item.TryGetProperty(colName, out JsonElement value))
            {
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
}
