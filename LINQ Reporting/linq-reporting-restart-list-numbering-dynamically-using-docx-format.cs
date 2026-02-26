using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list definition that will be reused for each group.
        List listDef = doc.Lists.Add(ListTemplate.NumberDefault);
        // Enable restarting the list at each new section (requires OOXML compliance > Ecma376).
        listDef.IsRestartAtEachSection = true;

        // Sample data to be reported via LINQ.
        var items = new[]
        {
            new { Category = "Fruits",      Name = "Apple"  },
            new { Category = "Fruits",      Name = "Banana" },
            new { Category = "Fruits",      Name = "Cherry" },
            new { Category = "Vegetables",  Name = "Carrot" },
            new { Category = "Vegetables",  Name = "Lettuce"},
            new { Category = "Vegetables",  Name = "Pepper" }
        };

        // Group items by Category and output each group as a separate list.
        foreach (var group in items.GroupBy(i => i.Category))
        {
            // Write the group heading.
            builder.Writeln(group.Key);

            // Start list formatting using the previously defined list.
            builder.ListFormat.List = listDef;

            // Add each item in the current group as a list entry.
            foreach (var item in group)
                builder.Writeln(item.Name);

            // End list formatting for this group.
            builder.ListFormat.RemoveNumbers();

            // Insert a section break so the next group's list restarts from 1.
            builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the document with a compliance level that writes IsRestartAtEachSection to DOCX.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("RestartListNumbering.docx", saveOptions);
    }
}
