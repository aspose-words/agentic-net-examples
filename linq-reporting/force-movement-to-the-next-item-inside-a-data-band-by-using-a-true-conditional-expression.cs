using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingTrueConditional
{
    public class Item
    {
        public int Index { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words if needed
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // Create template document with LINQ Reporting tags
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report generated with LINQ Reporting");
            builder.Writeln("<<foreach [item in Items]>>");
            // Output item details
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            // True conditional expression to force movement to next item (no additional content)
            builder.Writeln("<<if [true]>> <<[item.Name]>> <</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(templatePath);

            // Load the template for reporting
            Document doc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Build the report
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            doc.Save(outputPath);
        }
    }
}
