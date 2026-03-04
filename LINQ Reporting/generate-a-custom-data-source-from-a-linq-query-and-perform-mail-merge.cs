using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    // Simple data entity that matches the merge fields in the template.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
    }

    // Custom data source that wraps an IEnumerable<Customer> and implements IMailMergeDataSource.
    public class LinqCustomerDataSource : IMailMergeDataSource
    {
        private readonly IList<Customer> _customers;
        private int _recordIndex = -1; // Position before the first record.

        public LinqCustomerDataSource(IEnumerable<Customer> customers)
        {
            // Materialize the query to allow multiple passes.
            _customers = customers.ToList();
        }

        // Name of the data source (used for mail‑merge regions, not required here).
        public string TableName => "Customer";

        // Returns the value for a given field name.
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

        // Moves to the next record; returns false when the end is reached.
        public bool MoveNext()
        {
            if (!IsEof) _recordIndex++;
            return !IsEof;
        }

        // No child data sources are used in this simple example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    class Program
    {
        static void Main()
        {
            // Load the DOT template that contains MERGEFIELD tags (e.g., MERGEFIELD FullName, MERGEFIELD Address).
            Document template = new Document("Template.dot");

            // Sample collection of customers.
            List<Customer> allCustomers = new List<Customer>
            {
                new Customer { FullName = "Thomas Hardy", Address = "120 Hanover Sq., London" },
                new Customer { FullName = "Paolo Accorti", Address = "Via Monte Bianco 34, Torino" },
                new Customer { FullName = "John Doe", Address = "123 Main St., New York" }
            };

            // LINQ query to select only customers whose address contains "Torino".
            var filteredCustomers = allCustomers.Where(c => c.Address.Contains("Torino"));

            // Wrap the LINQ result in the custom mail‑merge data source.
            IMailMergeDataSource dataSource = new LinqCustomerDataSource(filteredCustomers);

            // Perform the mail merge using the custom data source.
            template.MailMerge.Execute(dataSource);

            // Save the merged document to a new file.
            template.Save("MergedResult.docx");
        }
    }
}
