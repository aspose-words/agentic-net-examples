using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model that matches the XML structure.
    public class Order
    {
        public int Id { get; set; }
        public string CustomerName { get; set; } = "";
        public DateTime Date { get; set; }
        public decimal Amount { get; set; }

        // Helper property that returns the amount formatted as currency.
        public string FormattedAmount => $"${Amount:N2}";
    }

    // Wrapper class that exposes the collection to the template.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample XML data file (kept for demonstration purposes).
            const string xmlFileName = "Orders.xml";
            GenerateLargeXmlData(xmlFileName, 1000);

            // Build the in‑memory data model that will be used by the reporting engine.
            ReportModel model = new();
            PopulateModelFromXml(xmlFileName, model);

            // Create the template document programmatically.
            const string templateFileName = "Template.docx";
            CreateTemplate(templateFileName);

            // Load the template.
            Document template = new Document(templateFileName);

            // Enable reflection optimization for better performance.
            ReportingEngine.UseReflectionOptimization = true;

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its members can be accessed in the template.
            engine.KnownTypes.Add(typeof(Order));

            // Build the report using the model. No data source name is required because the template
            // references the collection directly (<<foreach [order in Orders]>>).
            engine.BuildReport(template, model);

            // Save the generated report.
            const string outputFileName = "Report.docx";
            template.Save(outputFileName);
        }

        // Generates an XML file with a specified number of Order elements.
        private static void GenerateLargeXmlData(string filePath, int count)
        {
            XElement root = new("Orders");
            Random rnd = new();

            for (int i = 1; i <= count; i++)
            {
                Order order = new()
                {
                    Id = i,
                    CustomerName = $"Customer {i}",
                    Date = DateTime.Today.AddDays(-i),
                    Amount = (decimal)(rnd.NextDouble() * 1000 + 100)
                };

                XElement orderElement = new("Order",
                    new XElement("Id", order.Id),
                    new XElement("CustomerName", order.CustomerName),
                    new XElement("Date", order.Date.ToString("yyyy-MM-dd")),
                    new XElement("Amount", order.Amount));

                root.Add(orderElement);
            }

            XDocument doc = new(root);
            doc.Save(filePath);
        }

        // Populates the ReportModel from the generated XML file.
        private static void PopulateModelFromXml(string xmlPath, ReportModel model)
        {
            XDocument doc = XDocument.Load(xmlPath);
            foreach (XElement elem in doc.Root!.Elements("Order"))
            {
                Order order = new()
                {
                    Id = (int)elem.Element("Id")!,
                    CustomerName = (string)elem.Element("CustomerName")!,
                    Date = DateTime.Parse((string)elem.Element("Date")!),
                    Amount = (decimal)elem.Element("Amount")!
                };
                model.Orders.Add(order);
            }
        }

        // Creates a simple Word template containing LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header
            builder.Writeln("Order Report");
            builder.Writeln("------------------------------");

            // Begin foreach loop over orders.
            builder.Writeln("<<foreach [order in Orders]>>");

            // Table header
            builder.Writeln("Id\tCustomer\tDate\tAmount");
            builder.Writeln("------------------------------");

            // Table row with data fields.
            builder.Writeln(
                "<<[order.Id]>>\t" +
                "<<[order.CustomerName]>>\t" +
                "<<[order.Date]>>\t" +
                "<<[order.FormattedAmount]>>");

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            doc.Save(filePath);
        }
    }
}
