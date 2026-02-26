using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load a DOCM template that contains a list.
        Document doc = new Document("Template.docm");

        // After mail merge, restart list numbering at each new section.
        doc.MailMerge.RestartListsAtEachSection = true;

        // Create a data source using LINQ.
        var data = Enumerable.Range(1, 10)
            .Select(i => new
            {
                Item = $"Item {i}",
                // Alternate sections to demonstrate restarting.
                Section = i % 2 == 0 ? "EvenSection" : "OddSection"
            })
            .ToList();

        // Convert the LINQ result to a DataTable for MailMerge.
        DataTable table = new DataTable("Items");
        table.Columns.Add("Item");
        table.Columns.Add("Section");
        foreach (var row in data)
            table.Rows.Add(row.Item, row.Section);

        // Perform mail merge with regions. The template should contain a MERGEFIELD named "Item"
        // inside a table or list, and a SECTIONBREAK field to separate sections if needed.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the merged document as DOCM.
        doc.Save("MergedResult.docm", SaveFormat.Docm);
    }
}
