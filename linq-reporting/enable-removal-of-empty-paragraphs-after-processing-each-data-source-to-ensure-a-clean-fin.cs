using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingEmptyParagraphRemoval
{
    // Sample data models
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Order
    {
        public int Id { get; set; }
    }

    // Wrapper class for the root object used in the template
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document programmatically
            string templatePath = Path.Combine(outputDir, "template.docx");
            CreateTemplate(templatePath);

            // 2. Load the template
            Document doc = new Document(templatePath);

            // 3. Prepare sample data
            ReportModel model = new()
            {
                Customers = new()
                {
                    new Customer { Name = "Alice" },
                    new Customer { Name = "Bob" },
                    new Customer { Name = "Charlie" }
                },
                Orders = new()
                {
                    new Order { Id = 101 },
                    new Order { Id = 102 }
                }
            };

            // 4. Build report for the first data source (Customers)
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");
            RemoveEmptyParagraphs(doc); // Clean up after first processing

            // 5. Build report for the second data source (Orders)
            // The same document already contains the Orders section tags,
            // so we can run BuildReport again with the same model.
            engine.BuildReport(doc, model, "model");
            RemoveEmptyParagraphs(doc); // Clean up after second processing

            // 6. Save the final document
            string outputPath = Path.Combine(outputDir, "final_report.docx");
            doc.Save(outputPath);
        }

        // Creates a simple template with LINQ Reporting tags for Customers and Orders
        private static void CreateTemplate(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Customers section
            builder.Writeln("Customers:");
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("Name: <<[c.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln(""); // Intentional empty paragraph that may become redundant

            // Orders section
            builder.Writeln("Orders:");
            builder.Writeln("<<foreach [o in Orders]>>");
            builder.Writeln("Order Id: <<[o.Id]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln(""); // Another possible empty paragraph

            doc.Save(path);
        }

        // Removes paragraphs that contain no visible text (including whitespace)
        private static void RemoveEmptyParagraphs(Document doc)
        {
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            for (int i = paragraphs.Count - 1; i >= 0; i--)
            {
                Paragraph para = (Paragraph)paragraphs[i];
                // Paragraph.GetText() returns the text plus a paragraph break; trim removes whitespace and the break.
                if (string.IsNullOrWhiteSpace(para.GetText()))
                {
                    para.Remove();
                }
            }
        }
    }
}
