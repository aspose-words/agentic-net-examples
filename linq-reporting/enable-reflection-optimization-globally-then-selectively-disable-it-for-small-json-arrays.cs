using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingReflectionOptimization
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider required by Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Enable reflection optimization globally.
            ReportingEngine.UseReflectionOptimization = true;

            // Create sample JSON data files.
            string largeJsonPath = "large.json";
            string smallJsonPath = "small.json";
            CreateSampleJsonFiles(largeJsonPath, smallJsonPath);

            // Create and save the template document.
            string templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Load the template for the large data report.
            Document largeTemplate = new Document(templatePath);
            var largeData = LoadJsonData(largeJsonPath);
            var largeEngine = new ReportingEngine();
            largeEngine.BuildReport(largeTemplate, largeData, "data");
            largeTemplate.Save("LargeReport.docx");

            // Disable reflection optimization for the small data report.
            ReportingEngine.UseReflectionOptimization = false;

            // Load a fresh copy of the template for the small data report.
            Document smallTemplate = new Document(templatePath);
            var smallData = LoadJsonData(smallJsonPath);
            var smallEngine = new ReportingEngine();
            smallEngine.BuildReport(smallTemplate, smallData, "data");
            smallTemplate.Save("SmallReport.docx");
        }

        // Model classes matching the JSON structure.
        public class DataModel
        {
            public List<Item> Items { get; set; } = new();
        }

        public class Item
        {
            public string Name { get; set; } = "";
            public int Quantity { get; set; }
        }

        // Generates sample JSON files for large and small data sets.
        private static void CreateSampleJsonFiles(string largePath, string smallPath)
        {
            var largeModel = new DataModel();
            for (int i = 1; i <= 1000; i++)
            {
                largeModel.Items.Add(new Item { Name = $"Product {i}", Quantity = i });
            }
            File.WriteAllText(largePath, JsonConvert.SerializeObject(largeModel));

            var smallModel = new DataModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Quantity = 5 },
                    new Item { Name = "Banana", Quantity = 3 }
                }
            };
            File.WriteAllText(smallPath, JsonConvert.SerializeObject(smallModel));
        }

        // Loads JSON data into a strongly typed object for reporting.
        private static object LoadJsonData(string path)
        {
            string json = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<DataModel>(json)!;
        }

        // Creates a simple Word template containing LINQ Reporting tags.
        private static void CreateTemplate(string path)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Name]>> : <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");
            doc.Save(path);
        }
    }
}
