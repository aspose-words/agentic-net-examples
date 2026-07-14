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

        // Name of the data source (used only for mail‑merge regions).
        public string TableName => "Customer";

        // Moves to the next record. Returns false when the end of the collection is reached.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;

            return !IsEof;
        }

        // Retrieves the value for a given field name.
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
                    return false;
            }
        }

        // No child data sources are used in this simple example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Build a mail‑merge source document with the required MERGEFIELDs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("MERGEFIELD FullName");
            builder.InsertParagraph();
            builder.InsertField("MERGEFIELD Address");
            builder.InsertParagraph();

            // 2. Prepare a collection of Customer objects that will be merged.
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                new Customer("John Doe", "123 Main St., Anytown")
            };

            // 3. Wrap the collection in the custom data source.
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

            // 4. Execute the mail merge. This will generate a single document containing
            //    the merged content for each Customer in the collection.
            doc.MailMerge.Execute(dataSource);

            // 5. Save the merged document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCustomers.docx");
            doc.Save(outputPath);
        }
    }
}
