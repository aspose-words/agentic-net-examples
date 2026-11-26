using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalDefault
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data
            var model = new ReportModel();
            model.Items.Add(new Item { Name = "Alpha", Type = "A" });
            model.Items.Add(new Item { Name = "Beta", Type = "B" });
            model.Items.Add(new Item { Name = "Gamma", Type = "C" }); // No specific condition

            // Create template document
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Write a heading
            builder.Writeln("Items Report");
            builder.Writeln();

            // Begin foreach over Items
            builder.Writeln("<<foreach [item in Items]>>");

            // Conditional blocks with a default case
            builder.Writeln("<<if [item.Type == \"A\"]>>");
            builder.Writeln("Item: <<[item.Name]>> (Type A)");
            builder.Writeln("<</if>>");

            builder.Writeln("<<if [item.Type == \"B\"]>>");
            builder.Writeln("Item: <<[item.Name]>> (Type B)");
            builder.Writeln("<</if>>");

            // Default case when none of the above conditions are true
            builder.Writeln("<<if [!(item.Type == \"A\" || item.Type == \"B\")]>><<textColor [\"Red\"]>>");
            builder.Writeln("Item: <<[item.Name]>> (Other Type)");
            builder.Writeln("<</textColor>><</if>>");

            // End foreach
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            var outputPath = "output.docx";
            reportDoc.Save(outputPath);
        }
    }
}
