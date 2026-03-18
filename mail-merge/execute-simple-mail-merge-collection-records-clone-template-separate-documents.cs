using System;
using System.Collections.Generic;
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
            // Path to the template document that contains MERGEFIELDs:
            //   MERGEFIELD FullName
            //   MERGEFIELD Address
            string templatePath = @"C:\Templates\CustomerTemplate.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare a collection of records.
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
            };

            // For each record create a separate document.
            for (int i = 0; i < customers.Count; i++)
            {
                // Clone the template so that each document starts from the same base.
                Document doc = (Document)template.Clone();

                // Prepare field names and values for the current record.
                string[] fieldNames = { "FullName", "Address" };
                object[] fieldValues = { customers[i].FullName, customers[i].Address };

                // Perform a simple mail merge for a single record.
                // This uses the MailMerge.Execute(string[], object[]) overload as required.
                doc.MailMerge.Execute(fieldNames, fieldValues);

                // Save the merged document. Each file gets a unique name.
                string outputPath = $@"C:\Output\Customer_{i + 1}.docx";
                doc.Save(outputPath);
            }
        }
    }
}
