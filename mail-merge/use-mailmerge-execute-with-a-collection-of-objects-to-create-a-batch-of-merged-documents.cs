using System;
using System.Collections.Generic;
using System.IO;
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

    // Custom data source that implements Aspose.Words.IMailMergeDataSource.
    // This allows the MailMerge engine to retrieve values from the collection of Customer objects.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<Customer> _customers;
        private int _recordIndex = -1; // Position before the first record.

        public CustomerMailMergeDataSource(List<Customer> customers)
        {
            _customers = customers ?? throw new ArgumentNullException(nameof(customers));
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
                    return false; // Field not found.
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
            // 1. Create a blank document and add MERGEFIELDs that correspond to the data fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert fields for the customer's full name and address.
            builder.InsertField("MERGEFIELD FullName");
            builder.Writeln(); // New line after the name.
            builder.InsertField("MERGEFIELD Address");
            builder.Writeln(); // Ensure the document ends with a line break.

            // 2. Prepare a collection of Customer objects that will serve as the data source.
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"),
                new Customer("John Doe", "123 Main St., Anytown")
            };

            // 3. Wrap the collection in the custom data source implementation.
            CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

            // 4. Execute the mail merge. This will generate a merged document for each record in the collection.
            doc.MailMerge.Execute(dataSource);

            // 5. Save the resulting document to the file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedCustomers.docx");
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Mail merge completed. Document saved to: {outputPath}");
        }
    }
}
