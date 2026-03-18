using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsTemplateDemo
{
    // Simple data class that will be used in the data source.
    public class MyItem
    {
        public string Name { get; set; }
        public int Value { get; set; }
    }

    // Helper class that contains a static method which will be used as a delegate
    // inside the template's Select call. The method formats a MyItem instance.
    public static class TemplateHelpers
    {
        // This method matches the delegate signature expected by LINQ's Select.
        public static string FormatItem(MyItem item)
        {
            // Example formatting: "Name: Value"
            return $"{item.Name}: {item.Value}";
        }
    }

    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a template document in memory.
            // -------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The template uses the Reporting Engine syntax.
            // <<foreach [item in Items.Select(FormatItem)]>> iterates over the collection
            // and applies the custom delegate (static method) FormatItem to each element.
            builder.Writeln("<<foreach [item in Items.Select(FormatItem)]>>");
            builder.Writeln("- <<item>>"); // 'item' now holds the formatted string.
            builder.Writeln("<</foreach>>");

            // -------------------------------------------------
            // 2. Prepare the data source.
            // -------------------------------------------------
            var dataSource = new
            {
                Items = new List<MyItem>
                {
                    new MyItem { Name = "Alpha",   Value = 100 },
                    new MyItem { Name = "Beta",    Value = 200 },
                    new MyItem { Name = "Gamma",   Value = 300 }
                }
            };

            // -------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register the helper type so that the template can resolve the static method.
            engine.KnownTypes.Add(typeof(TemplateHelpers));

            // -------------------------------------------------
            // 4. Build the report using the template and data source.
            // -------------------------------------------------
            // The third argument ("data") is the name used to reference the data source
            // inside the template (e.g., Items refers to data.Items).
            engine.BuildReport(template, dataSource, "data");

            // -------------------------------------------------
            // 5. Save the generated document.
            // -------------------------------------------------
            template.Save("FormattedReport.docx");
        }
    }
}
