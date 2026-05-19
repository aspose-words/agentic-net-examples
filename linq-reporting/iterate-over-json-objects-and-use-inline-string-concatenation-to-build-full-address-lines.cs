using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple address model – used only for JSON serialization, not directly by the report.
    public class Address
    {
        public string Street { get; set; } = "";
        public string City { get; set; } = "";
        public string Country { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample JSON data (an array of address objects).
            var sampleAddresses = new List<Address>
            {
                new Address { Street = "221B Baker Street", City = "London", Country = "UK" },
                new Address { Street = "1600 Pennsylvania Ave NW", City = "Washington", Country = "USA" },
                new Address { Street = "1 Infinite Loop", City = "Cupertino", Country = "USA" }
            };

            // Serialize to JSON and write to a local file.
            string jsonPath = "addresses.json";
            File.WriteAllText(jsonPath, System.Text.Json.JsonSerializer.Serialize(sampleAddresses));

            // Create a Word document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a heading.
            builder.Writeln("Address Report");
            builder.Writeln("----------------");

            // Begin a foreach loop over the JSON array (named "addresses").
            builder.Writeln("<<foreach [addr in addresses]>>");
            // Build a full address line using inline string concatenation.
            builder.Writeln("<<[addr.Street + \", \" + addr.City + \", \" + addr.Country]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Load the JSON data as a data source for the reporting engine.
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "addresses");

            // Save the generated report.
            doc.Save("AddressReport.docx");

            // Clean up temporary JSON file.
            File.Delete(jsonPath);
        }
    }
}
