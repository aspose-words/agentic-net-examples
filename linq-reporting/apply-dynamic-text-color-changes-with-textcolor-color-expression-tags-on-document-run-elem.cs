using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Text = "Success", Color = "Green" },
                    new Item { Text = "Warning", Color = "Orange" },
                    new Item { Text = "Error",   Color = "Red" }
                }
            };

            // Create a template document containing textColor tags.
            const string templateFile = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<textColor [item.Color]>>[item.Text]<</textColor>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templateFile);

            // Load the template and generate the report.
            var reportDoc = new Document(templateFile);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");
            reportDoc.Save("ReportOutput.docx");
        }
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Text { get; set; } = "";
        public string Color { get; set; } = "";
    }
}
