using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

class CsvNestedBullets
{
    static void Main()
    {
        // Sample CSV data (Category,Item). In a real scenario you could read from a file.
        string[] csvLines = new[]
        {
            "Category,Item",          // header
            "Fruits,Apple",
            "Fruits,Banana",
            "Fruits,Orange",
            "Vegetables,Carrot",
            "Vegetables,Potato",
            "Vegetables,Tomato"
        };

        // Skip the header and any empty lines.
        var dataLines = csvLines.Skip(1)
                                .Where(l => !string.IsNullOrWhiteSpace(l));

        // Parse each line into a tuple (Category, Item) and group by Category.
        var groups = dataLines
            .Select(line => line.Split(','))
            .Where(parts => parts.Length >= 2)
            .GroupBy(parts => parts[0].Trim());

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a multilevel bulleted list based on the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Iterate over each category group and write nested bullet points.
        foreach (var group in groups)
        {
            // Write the category as a top‑level bullet (level 0).
            builder.ListFormat.List = bulletList;
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln(group.Key);

            // Write each item under the category as a second‑level bullet (level 1).
            foreach (var parts in group)
            {
                builder.ListFormat.ListLevelNumber = 1;
                builder.Writeln(parts[1].Trim());
            }

            // Add an empty paragraph to separate groups (optional).
            builder.Writeln();
        }

        // End list formatting for any subsequent paragraphs.
        builder.ListFormat.List = null;

        // Save the document to the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "NestedBullets.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
