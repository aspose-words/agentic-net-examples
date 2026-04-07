using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity representing an item in the JSON array.
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
    }

    // Wrapper model passed to the reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for environments that require it.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for temporary files.
            string dataFile = "data.json";
            string templateFile = "template.docx";
            string outputFile = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create sample JSON data.
            // -----------------------------------------------------------------
            string jsonContent = @"[
                { ""Name"": ""Alpha"",   ""Category"": ""Important"" },
                { ""Name"": ""Beta"",    ""Category"": ""Other"" },
                { ""Name"": ""Gamma"",   ""Category"": ""Important"" },
                { ""Name"": ""Delta"",   ""Category"": ""Other"" },
                { ""Name"": ""Epsilon"", ""Category"": ""Important"" }
            ]";

            File.WriteAllText(dataFile, jsonContent);

            // -----------------------------------------------------------------
            // 2. Load JSON, filter the array, and prepare the model.
            // -----------------------------------------------------------------
            var allItems = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(dataFile)) ?? new();
            var filteredItems = allItems.Where(i => i.Category == "Important").ToList();

            var model = new ReportModel { Items = filteredItems };

            // -----------------------------------------------------------------
            // 3. Build the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Apply a bulleted list style to the following paragraphs.
            builder.ListFormat.List = templateDoc.Lists.Add(ListTemplate.BulletDefault);

            // LINQ Reporting tags: foreach over Items and output each Name.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // End the list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the template to disk (required before BuildReport).
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // 4. Load the template and generate the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templateFile);
            var engine = new ReportingEngine();

            // Build the report using the model; the root name must match the tag references.
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document.
            reportDoc.Save(outputFile);
        }
    }
}
