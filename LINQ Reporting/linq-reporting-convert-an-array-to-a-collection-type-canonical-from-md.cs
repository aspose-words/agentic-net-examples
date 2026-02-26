using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document with a foreach block that expects a collection named "items"
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report:");
        // The syntax <<foreach [items]>><<[Name]>>...<<</foreach>> iterates over each element in the collection
        builder.Writeln("<<foreach [items]>><<[Name]>> <<</foreach>>");

        // Prepare the data source as an array of objects
        DataItem[] array = new DataItem[]
        {
            new DataItem { Name = "Alice", Age = 30 },
            new DataItem { Name = "Bob", Age = 25 },
            new DataItem { Name = "Charlie", Age = 35 }
        };

        // Convert the array to a List<DataItem>, which is the canonical collection type recognized by ReportingEngine
        List<DataItem> list = array.ToList();

        // Build the report using the list as the data source.
        // The third argument ("items") must match the collection name used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, list, "items");

        // Save the populated document.
        template.Save("Report.docx");
    }

    // Simple POCO class used as the data source.
    public class DataItem
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
