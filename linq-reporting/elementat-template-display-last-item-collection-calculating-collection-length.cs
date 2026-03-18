using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting; // Reporting Engine namespace

class Program
{
    static void Main()
    {
        // Load the Word template that contains a reporting field using ElementAt.
        // The template should have a field like:
        // <<Items.ElementAt(Items.Count-1)>>
        // which evaluates to the last element of the collection.
        Document doc = new Document("Template.docx");

        // Prepare a collection that will be used in the template.
        List<string> items = new List<string>
        {
            "Apple",
            "Banana",
            "Cherry"
        };

        // Wrap the collection in an anonymous object so the reporting engine can access it.
        var dataSource = new { Items = items };

        // Build the report. The root name ("Items") matches the property name in the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "Items");

        // Save the populated document.
        doc.Save("Result.docx");
    }
}
