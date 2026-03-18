using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeBatchExample
{
    // Simple data entity that will be merged into the document.
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
        private readonly IList<Customer> _customers;
        private int _recordIndex = -1; // Position before the first record.

        public CustomerMailMergeDataSource(IList<Customer> customers)
        {
            _customers = customers;
        }

        // Name of the data source (used only for mail merge with regions).
        public string TableName => "Customer";

        // Moves to the next record. Returns false when the end of the collection is reached.
        public bool MoveNext()
        {
            if (_recordIndex < _customers.Count - 1)
            {
                _recordIndex++;
                return true;
            }
            return false;
        }

        // Returns the value for a given field name from the current record.
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
    }

    class Program
    {
        static void Main()
        {
            // Create a blank document and add two MERGEFIELDs.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare a collection of Customer objects to be merged.
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                new Customer("John Doe", "123 Main St., New York")
            };

            // Wrap the collection in the custom data source.
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

            // Perform the mail merge. This will generate a separate document for each record.
            doc.MailMerge.Execute(dataSource);

            // Save the merged document.
            doc.Save("MergedCustomers.docx");
        }
    }
}
