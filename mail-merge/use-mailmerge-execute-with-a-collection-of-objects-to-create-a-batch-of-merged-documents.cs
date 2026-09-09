using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeBatchExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank document and add merge fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD FullName");
            builder.InsertParagraph();
            builder.InsertField("MERGEFIELD Address");
            builder.InsertParagraph();

            // Prepare a collection of data objects.
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
            };

            // Wrap the collection in a custom mail‑merge data source.
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

            // Execute the mail merge. All records will be merged into the same document,
            // each record producing a new copy of the document content.
            doc.MailMerge.Execute(dataSource);

            // Save the merged document.
            doc.Save("MergedCustomers.docx");
        }
    }

    // Simple data entity used for the mail merge.
    public class Customer
    {
        public Customer(string fullName, string address)
        {
            FullName = fullName;
            Address = address;
        }

        public string FullName { get; }
        public string Address { get; }
    }

    // Custom data source that implements IMailMergeDataSource.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly IList<Customer> _customers;
        private int _recordIndex = -1;

        public CustomerMailMergeDataSource(IList<Customer> customers)
        {
            _customers = customers;
        }

        // Name of the data source (used only for regions).
        public string TableName => "Customer";

        // Move to the next record in the collection.
        public bool MoveNext()
        {
            if (_recordIndex < _customers.Count - 1)
            {
                _recordIndex++;
                return true;
            }
            return false;
        }

        // Return the value for the requested field name.
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

        // No child data sources are used in this example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;
    }
}
