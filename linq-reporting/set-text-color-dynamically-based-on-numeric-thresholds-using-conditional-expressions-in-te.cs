using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Score { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create the template document with LINQ Reporting tags
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");

            // Use a conditional expression to choose the text color based on Score
            // Green for >=80, Orange for >=50, otherwise Red
            builder.Writeln(
                "<<textColor [item.Score >= 80 ? \"Green\" : item.Score >= 50 ? \"Orange\" : \"Red\"]>>" +
                "Name: <<[item.Name]>>, Score: <<[item.Score]>>" +
                " <</textColor>>");

            // End the foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Step 2: Prepare sample data
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Alice", Score = 92 },
                    new Item { Name = "Bob",   Score = 76 },
                    new Item { Name = "Carol", Score = 43 }
                }
            };

            // Step 3: Load the template and build the report
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Step 4: Save the generated report
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
