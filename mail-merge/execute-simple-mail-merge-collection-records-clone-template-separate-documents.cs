using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    // Simple data entity.
    public class Customer
    {
        public Customer(string fullName, string address)
        {
            FullName = fullName;
            Address = address;
        }

        public string FullName { get; set; }
        public string Address { get; set; }
    }

    public class Program
    {
        // Entry point.
        public static void Main()
        {
            // Create a simple template document in memory with MERGEFIELDs.
            Document template = CreateTemplateDocument();

            // Prepare a collection of records.
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
            };

            // Ensure output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // For each record create a separate document.
            for (int i = 0; i < customers.Count; i++)
            {
                // Clone the template so that each document starts from the same base.
                Document doc = (Document)template.Clone();

                // Prepare field names and values for the current record.
                string[] fieldNames = { "FullName", "Address" };
                object[] fieldValues = { customers[i].FullName, customers[i].Address };

                // Perform a simple mail merge for a single record.
                doc.MailMerge.Execute(fieldNames, fieldValues);

                // Save the merged document. Each file gets a unique name.
                string outputPath = Path.Combine(outputDir, $"Customer_{i + 1}.docx");
                doc.Save(outputPath);
                Console.WriteLine($"Saved merged document to: {outputPath}");
            }
        }

        // Helper method to create a template document with merge fields.
        private static Document CreateTemplateDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Customer Information");
            builder.Writeln();

            builder.InsertField("MERGEFIELD FullName");
            builder.Writeln();
            builder.InsertField("MERGEFIELD Address");
            builder.Writeln();

            return doc;
        }
    }
}
