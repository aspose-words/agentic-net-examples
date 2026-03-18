using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data to be grouped.
        var data = new[]
        {
            new { Category = "Fruits", Item = "Apple" },
            new { Category = "Fruits", Item = "Banana" },
            new { Category = "Fruits", Item = "Cherry" },
            new { Category = "Vegetables", Item = "Carrot" },
            new { Category = "Vegetables", Item = "Broccoli" },
            new { Category = "Grains", Item = "Rice" },
            new { Category = "Grains", Item = "Wheat" }
        };

        // Create a bulleted list based on the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList; // Apply the list to the builder.

        // Group the data by Category.
        var groups = data.GroupBy(d => d.Category);

        foreach (var group in groups)
        {
            // Write the group heading as a top‑level bullet (level 0).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln(group.Key);

            // Increase the list level to create nested bullets for the items.
            builder.ListFormat.ListIndent(); // Now at level 1.

            foreach (var entry in group)
            {
                builder.Writeln(entry.Item);
            }

            // Return to the previous level before processing the next group.
            builder.ListFormat.ListOutdent(); // Back to level 0.
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("NestedBulletedGroups.docx");
    }
}
