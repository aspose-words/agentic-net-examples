using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper object for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("LINQ Reporting – Force next item with a true condition");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Index]>> – <<[item.Name]>>");
            // The true condition forces the engine to move to the next item.
            builder.Writeln("<<if [true]>><<next>><</if>>");
            // This line will be skipped because of the <<next>> tag above.
            builder.Writeln("This line is skipped for item <<[item.Index]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for report generation.
            Document reportDocument = new Document(templatePath);

            // Prepare sample data.
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(reportDocument, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            reportDocument.Save(outputPath);
        }
    }
}
