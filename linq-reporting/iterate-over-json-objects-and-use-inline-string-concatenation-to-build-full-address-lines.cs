using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model representing a person with address fields.
    public class Person
    {
        public string Street { get; set; } = string.Empty;
        public string City   { get; set; } = string.Empty;
        public string State  { get; set; } = string.Empty;
        public string Zip    { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Prepare sample JSON data.
            // -----------------------------------------------------------------
            var people = new List<Person>
            {
                new Person { Street = "123 Main St", City = "Springfield", State = "IL", Zip = "62704" },
                new Person { Street = "456 Oak Ave",  City = "Metropolis",  State = "NY", Zip = "10001" },
                new Person { Street = "789 Pine Rd",  City = "Gotham",      State = "NJ", Zip = "07097" }
            };

            const string jsonFile = "people.json";
            File.WriteAllText(jsonFile, JsonConvert.SerializeObject(people, Formatting.Indented));

            // -----------------------------------------------------------------
            // 2. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            const string templateFile = "AddressTemplate.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the root collection "persons".
            builder.Writeln("<<foreach [person in persons]>>");

            // Build a full address line using inline string concatenation.
            // Note the escaped double quotes for C# string literals.
            builder.Writeln(
                "<<[person.Street + \", \" + person.City + \", \" + person.State + \" \" + person.Zip]>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the JSON data source.
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templateFile);
            var jsonDataSource = new JsonDataSource(jsonFile);

            // -----------------------------------------------------------------
            // 4. Build the report using ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // The root object name is "persons" (matches the JSON array).
            engine.BuildReport(loadedTemplate, jsonDataSource, "persons");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputFile = "AddressReport.docx";
            loadedTemplate.Save(outputFile);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputFile)}");
        }
    }
}
