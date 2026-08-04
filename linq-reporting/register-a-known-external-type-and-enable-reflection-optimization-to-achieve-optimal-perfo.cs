using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data entity.
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public DateTime OrderDate { get; set; }
    }

    // Wrapper model that holds a collection of orders.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    // External type whose static members can be used inside the template.
    public static class Utils
    {
        // Formats a DateTime value; will be called from the template.
        public static string FormatDate(DateTime dt) => dt.ToString("yyyy-MM-dd");
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the working directory exists.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // 1. Create a sample XML data file.
            string xmlPath = Path.Combine(workDir, "Orders.xml");
            string xmlContent =
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderDate>2023-01-15T00:00:00</OrderDate>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderDate>2023-02-20T00:00:00</OrderDate>
    </Order>
    <Order>
        <CustomerName>Bob Johnson</CustomerName>
        <OrderDate>2023-03-05T00:00:00</OrderDate>
    </Order>
</Orders>";
            File.WriteAllText(xmlPath, xmlContent);

            // 2. Create the template document programmatically.
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert LINQ Reporting tags.
            builder.Writeln("<<foreach [order in model.Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Date: <<[Utils.FormatDate(order.OrderDate)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // 3. Load the template document.
            Document doc = new Document(templatePath);

            // 4. Load XML data into a list of Order objects.
            // For simplicity we parse the XML manually; in a real scenario the XML could be huge.
            var orders = new List<Order>();
            var xmlDoc = new System.Xml.XmlDocument();
            xmlDoc.Load(xmlPath);
            foreach (System.Xml.XmlNode node in xmlDoc.SelectNodes("//Order"))
            {
                var order = new Order
                {
                    CustomerName = node["CustomerName"]?.InnerText ?? "",
                    OrderDate = DateTime.Parse(node["OrderDate"]?.InnerText ?? DateTime.MinValue.ToString())
                };
                orders.Add(order);
            }

            // 5. Prepare the model.
            var model = new ReportModel { Orders = orders };

            // 6. Enable reflection optimization (static property).
            ReportingEngine.UseReflectionOptimization = true;

            // 7. Create the reporting engine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Utils));

            // 8. Build the report.
            engine.BuildReport(doc, model, "model");

            // 9. Save the generated report.
            string outputPath = Path.Combine(workDir, "Report.docx");
            doc.Save(outputPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine("Report generated at: " + outputPath);
        }
    }
}
