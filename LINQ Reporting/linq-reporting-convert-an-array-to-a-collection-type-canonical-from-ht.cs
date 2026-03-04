using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an HTML template that expects a collection named "items".
        // The template uses the foreach tag to iterate over the collection.
        string htmlTemplate = @"
<p>Items list:</p>
<<foreach [items]>>
    <<[Name]>>
<<endforeach>>
";
        builder.InsertHtml(htmlTemplate);

        // Prepare the data source: an array of Item objects.
        Item[] array = new Item[]
        {
            new Item { Name = "Apple" },
            new Item { Name = "Banana" },
            new Item { Name = "Cherry" }
        };

        // Convert the array to a List<T>, which is the canonical collection type
        // recognized by the LINQ Reporting Engine.
        List<Item> items = array.ToList();

        // Build the report. The data source is an anonymous object that exposes the
        // collection under the name "items" used in the template.
        ReportingEngine engine = new ReportingEngine();
        var dataSource = new { items = items };
        engine.BuildReport(doc, dataSource);

        // Save the resulting document.
        doc.Save("ReportFromArray.docx");
    }
}
