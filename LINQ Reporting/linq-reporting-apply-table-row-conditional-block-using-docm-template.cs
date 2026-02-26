using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; }
    public bool IsActive { get; set; }
}

public class ReportData
{
    public List<Item> Items { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Load the DOCM template that contains a table with a conditional block.
        // The template should have a row like:
        // <<foreach [in Items]>>
        //   <<if [IsActive]>>
        //     <<[Name]>>
        //   <<endif>>
        // <<endforeach>>
        Document doc = new Document("Template.docm");

        // Prepare the data source.
        var data = new ReportData
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", IsActive = true },
                new Item { Name = "Beta",  IsActive = false },
                new Item { Name = "Gamma", IsActive = true }
            }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the data source.
        // The third argument ("data") is the name used to reference the data source inside the template.
        engine.BuildReport(doc, data, "data");

        // Save the populated document.
        doc.Save("Result.docx");
    }
}
