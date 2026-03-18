using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Sample data to be grouped.
        var data = new[]
        {
            new { Category = "Fruits",      Item = "Apple"  },
            new { Category = "Fruits",      Item = "Banana" },
            new { Category = "Fruits",      Item = "Orange" },
            new { Category = "Vegetables", Item = "Carrot" },
            new { Category = "Vegetables", Item = "Broccoli" },
            new { Category = "Grains",     Item = "Rice" }
        };

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list based on the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Group the data by Category.
        var groups = data.GroupBy(d => d.Category);

        foreach (var group in groups)
        {
            // Start a new list for the current group.
            builder.ListFormat.List = bulletList;

            // Write the group title as a top‑level bullet (level 0).
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln(group.Key);

            // Increase the list level to create a nested bullet section.
            builder.ListFormat.ListIndent(); // now at level 1

            // Write each item of the group as a nested bullet.
            foreach (var entry in group)
            {
                builder.Writeln(entry.Item);
            }

            // Return to the previous level and end the list for this group.
            builder.ListFormat.ListOutdent(); // back to level 0
            builder.ListFormat.RemoveNumbers(); // closes the list

            // Add an empty paragraph to separate groups.
            builder.Writeln();
        }

        // Save the resulting document.
        doc.Save("GroupedBullets.docx");
    }
}
