using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Create first list (numbered) ----------
        // Add a numbered list to the document using a built‑in template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        // Assign the list to the builder.
        builder.ListFormat.List = numberedList;
        // Write list items.
        builder.Writeln("Numbered item 1");
        builder.Writeln("Numbered item 2");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // ---------- Create second list (bulleted) ----------
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.Writeln("Bullet item 1");
        builder.Writeln("Bullet item 2");
        builder.ListFormat.RemoveNumbers();

        // ---------- Apply a uniform style to all lists ----------
        // Iterate over every list in the document.
        foreach (List list in doc.Lists)
        {
            // Iterate over each level of the current list (max 9 levels).
            foreach (ListLevel level in list.ListLevels)
            {
                // Use a numeric format that produces decimal numbers.
                level.NumberFormat = "%1.";
                // Set a uniform font for the list level.
                level.Font.Name = "Times New Roman";
                // Ensure numbering starts at 1 for each level.
                level.StartAt = 1;
            }
        }

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
