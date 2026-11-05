using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace LinqReportingMultilingual
{
    // Root data model aligned with the template.
    public class ReportModel
    {
        public string TitleEn { get; set; } = "";
        public string TitleFr { get; set; } = "";
        public string TitleZh { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string NameEn { get; set; } = "";
        public string NameFr { get; set; } = "";
        public string NameZh { get; set; } = "";
        public decimal Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for full Unicode support.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for temporary files.
            string templatePath = "template.docx";
            string jsonPath = "data.json";
            string outputPath = "report.docx";

            // Create sample JSON data.
            var sampleModel = new ReportModel
            {
                TitleEn = "Product Catalog (English)",
                TitleFr = "Catalogue de Produits (Français)",
                TitleZh = "产品目录 (中文)",
                Items = new List<Item>
                {
                    new Item { NameEn = "Apple", NameFr = "Pomme", NameZh = "苹果", Price = 1.20m },
                    new Item { NameEn = "Banana", NameFr = "Banane", NameZh = "香蕉", Price = 0.80m },
                    new Item { NameEn = "Cherry", NameFr = "Cerise", NameZh = "樱桃", Price = 2.50m }
                }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleModel, Formatting.Indented));

            // -----------------------------------------------------------------
            // Build the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title section with multilingual fields.
            builder.Writeln("<<[model.TitleEn]>>");
            builder.Writeln("<<[model.TitleFr]>>");
            builder.Writeln("<<[model.TitleZh]>>");
            builder.Writeln();

            // Header table (static, not repeated).
            builder.Writeln("Items:");
            Table headerTable = builder.StartTable();
            builder.InsertCell(); builder.Writeln("Name (EN)");
            builder.InsertCell(); builder.Writeln("Name (FR)");
            builder.InsertCell(); builder.Writeln("Name (ZH)");
            builder.InsertCell(); builder.Writeln("Price");
            builder.EndRow();
            builder.EndTable();

            // Repeated rows – each iteration creates a one‑row table.
            builder.Writeln("<<foreach [item in Items]>>");
            Table dataTable = builder.StartTable();
            builder.InsertCell(); builder.Writeln("<<[item.NameEn]>>");
            builder.InsertCell(); builder.Writeln("<<[item.NameFr]>>");
            builder.InsertCell(); builder.Writeln("<<[item.NameZh]>>");
            builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load JSON data.
            // -----------------------------------------------------------------
            string jsonContent = File.ReadAllText(jsonPath);
            var model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
