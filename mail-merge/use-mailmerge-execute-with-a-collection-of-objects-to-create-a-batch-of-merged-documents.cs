using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeBatchExample
{
    // Simple data entity representing a customer.
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

    // Custom mail merge data source that wraps a collection of Customer objects.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<Customer> _customers;
        private int _recordIndex = -1; // Position before the first record.

        public CustomerMailMergeDataSource(List<Customer> customers)
        {
            _customers = customers;
        }

        // Name of the data source (used for mail‑merge regions, not needed here).
        public string TableName => "Customer";

        // Moves to the next record. Returns false when the end of the collection is reached.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;

            return !IsEof;
        }

        // Retrieves the value for a given field name from the current record.
        public bool GetValue(string fieldName, out object fieldValue)
        {
            switch (fieldName)
            {
                case "FullName":
                    fieldValue = _customers[_recordIndex].FullName;
                    return true;
                case "Address":
                    fieldValue = _customers[_recordIndex].Address;
                    return true;
                default:
                    fieldValue = null;
                    return false; // Field not found.
            }
        }

        // No child data sources are used in this example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and add merge fields for FullName and Address.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD FullName", "<FullName>");
            builder.Writeln(); // New line between fields.
            builder.InsertField("MERGEFIELD Address", "<Address>");

            // Prepare a collection of customers to merge.
            List<Customer> customers = new List<Customer>
            {
                new Customer("John Doe", "123 Main St, Anytown"),
                new Customer("Jane Smith", "456 Oak Ave, Othertown"),
                new Customer("Bob Johnson", "789 Pine Rd, Sometown")
            };

            // Wrap the collection in the custom data source.
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

            // Execute the mail merge. This will generate a merged document for each record.
            doc.MailMerge.Execute(dataSource);

            // Save the merged document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Mail merge completed. Document saved to: {outputPath}");
        }
    }
}
