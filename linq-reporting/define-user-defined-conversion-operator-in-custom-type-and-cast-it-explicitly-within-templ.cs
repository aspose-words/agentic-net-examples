using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConversionExample
{
    // Custom type with an explicit conversion operator to string.
    public class MyValue
    {
        public int Number { get; set; }

        // Explicit conversion to string.
        public static explicit operator string(MyValue value) => value.Number.ToString();
    }

    // Model class used in the report.
    public class Item
    {
        public MyValue Value { get; set; } = new MyValue();
    }

    // Public wrapper for the data source required by the ReportingEngine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var items = new List<Item>
            {
                new() { Value = new MyValue { Number = 10 } },
                new() { Value = new MyValue { Number = 20 } },
                new() { Value = new MyValue { Number = 30 } }
            };

            // Create the template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Sample Report");
            builder.Writeln("<<foreach [item in Items]>>");
            // Explicitly cast MyValue to string using the user‑defined conversion operator.
            builder.Writeln("Value: <<[(string)item.Value]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template for reporting.
            var template = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();

            var model = new ReportModel { Items = items };
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
