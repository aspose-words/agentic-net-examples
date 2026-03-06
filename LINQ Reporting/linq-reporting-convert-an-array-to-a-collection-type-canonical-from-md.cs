using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Sample data defined as an array.
        DataItem[] dataArray = new DataItem[]
        {
            new DataItem { Name = "Item1", Value = 10 },
            new DataItem { Name = "Item2", Value = 20 }
        };

        // Convert the array to a List<T>, which is the canonical collection type
        // expected by the Aspose.Words ReportingEngine.
        List<DataItem> dataList = dataArray.ToList();

        // Create a simple template document in memory.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a repeatable region that iterates over the collection "ds".
        // Inside the region we output the Name and Value fields.
        builder.Writeln("<<foreach [ds]>><<[Name]>>: <<[Value]>>\n<</foreach>>");

        // Build the report using the list as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataList, "ds");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Simple POCO class used as the data source.
public class DataItem
{
    public string Name { get; set; }
    public int Value { get; set; }
}
