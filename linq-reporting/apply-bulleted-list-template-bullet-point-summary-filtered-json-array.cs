using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListFromJson
{
    // Simple model matching the JSON structure.
    public class Item
    {
        public string Title { get; set; }
        public string Category { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Sample JSON data embedded directly in the source.
            const string jsonData = @"
            [
                { ""Title"": ""Prepare quarterly report"", ""Category"": ""Important"" },
                { ""Title"": ""Schedule team meeting"", ""Category"": ""Normal"" },
                { ""Title"": ""Update project roadmap"", ""Category"": ""Important"" },
                { ""Title"": ""Review pull requests"", ""Category"": ""Normal"" }
            ]";

            // Deserialize the JSON array.
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            List<Item> allItems = JsonSerializer.Deserialize<List<Item>>(jsonData, options);

            // Filter the items – keep only those with Category == ""Important"".
            List<Item> filtered = allItems.FindAll(i => i.Category == "Important");

            // ---------- Document lifecycle ----------
            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Create a DocumentBuilder for inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 3. Add a bulleted list based on the default bullet template.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // 4. Apply the list to the builder.
            builder.ListFormat.List = bulletList;
            builder.ListFormat.ListLevelNumber = 0; // Ensure we are on the first level.

            // 5. Write each filtered item as a bullet point.
            foreach (Item item in filtered)
            {
                builder.Writeln(item.Title);
            }

            // 6. End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // 7. Save the resulting document.
            doc.Save("BulletSummary.docx");
        }
    }
}
