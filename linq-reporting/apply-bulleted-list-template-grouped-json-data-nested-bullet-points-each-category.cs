using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the JSON file that contains categories and their items.
        // Expected format: { "Category1": ["Item1", "Item2"], "Category2": ["ItemA"] }
        string jsonPath = "data.json";

        // If the file does not exist, use a default JSON string.
        string jsonContent = File.Exists(jsonPath)
            ? File.ReadAllText(jsonPath)
            : "{\"Fruits\":[\"Apple\",\"Banana\"],\"Vegetables\":[\"Carrot\",\"Broccoli\"]}";

        // Deserialize the JSON into a dictionary where the key is the category
        // and the value is a list of items belonging to that category.
        var groupedData = JsonSerializer.Deserialize<Dictionary<string, List<string>>>(jsonContent)
                          ?? new Dictionary<string, List<string>>();

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list using the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Build nested bullet points: category as level 0, items as level 1.
        foreach (var kvp in groupedData)
        {
            // Write the category (top‑level bullet).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln(kvp.Key);

            // Increase the list level to write the items under the category.
            builder.ListFormat.ListIndent(); // level 1
            foreach (string item in kvp.Value)
            {
                builder.Writeln(item);
            }
            // Return to the previous level before processing the next category.
            builder.ListFormat.ListOutdent(); // back to level 0
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("NestedBullets.docx");
    }
}
