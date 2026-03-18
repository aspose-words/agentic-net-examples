using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace CustomMailMergeExample
{
    // Simple data entity.
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

    // Custom data source that implements both IMailMergeDataSource (required by Aspose.Words)
    // and IEnumerable<Customer> (allows streaming large data sets lazily).
    public class CustomerDataSource : IMailMergeDataSource, IEnumerable<Customer>
    {
        private readonly IEnumerable<Customer> _source;          // Original enumerable (could be lazy).
        private IEnumerator<Customer> _enumerator;               // Enumerator used by the mail merge engine.
        private Customer _current;                               // Holds the current record after MoveNext.

        public CustomerDataSource(IEnumerable<Customer> source)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            // Create a fresh enumerator for the mail merge engine.
            _enumerator = _source.GetEnumerator();
        }

        // IMailMergeDataSource implementation ---------------------------------

        // Name of the data source – must match the merge region name if regions are used.
        public string TableName => "Customer";

        // Advances to the next record. Stores the current record for GetValue.
        public bool MoveNext()
        {
            if (_enumerator.MoveNext())
            {
                _current = _enumerator.Current;
                return true;
            }
            return false;
        }

        // Returns the value for a given field name.
        public bool GetValue(string fieldName, out object fieldValue)
        {
            switch (fieldName)
            {
                case "FullName":
                    fieldValue = _current.FullName;
                    return true;
                case "Address":
                    fieldValue = _current.Address;
                    return true;
                default:
                    fieldValue = null;
                    return false;
            }
        }

        // No child data sources in this simple example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        // IEnumerable<Customer> implementation ----------------------------------

        public IEnumerator<Customer> GetEnumerator() => _source.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    class Program
    {
        // Example of a lazy enumerable that could stream millions of records.
        private static IEnumerable<Customer> GetLargeCustomerEnumerable()
        {
            // In a real scenario this could read from a database, file, etc.
            // Here we simulate with a simple loop.
            for (int i = 1; i <= 1000000; i++)
            {
                yield return new Customer($"Customer {i}", $"Address {i}");
            }
        }

        static void Main()
        {
            // Create a blank document and add merge fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(" MERGEFIELD FullName ");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD Address ");

            // Prepare the custom data source that streams the data.
            IEnumerable<Customer> customers = GetLargeCustomerEnumerable();
            CustomerDataSource dataSource = new CustomerDataSource(customers);

            // Execute mail merge using the custom data source.
            doc.MailMerge.Execute(dataSource);

            // Save the result.
            doc.Save("MergedCustomers.docx");
        }
    }
}
