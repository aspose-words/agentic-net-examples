using System;
using System.Collections;
using System.Collections.Generic;

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

    // Minimal IMailMergeDataSource definition (mirrors the Aspose interface used in the example).
    public interface IMailMergeDataSource
    {
        string TableName { get; }
        bool MoveNext();
        bool GetValue(string fieldName, out object fieldValue);
        IMailMergeDataSource GetChildDataSource(string tableName);
    }

    // Custom data source that implements both IMailMergeDataSource and IEnumerable<Customer>.
    public class CustomerDataSource : IMailMergeDataSource, IEnumerable<Customer>
    {
        private readonly IEnumerable<Customer> _source;
        private IEnumerator<Customer> _enumerator;
        private Customer _current;

        public CustomerDataSource(IEnumerable<Customer> source)
        {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            _enumerator = _source.GetEnumerator();
        }

        // IMailMergeDataSource implementation
        public string TableName => "Customer";

        public bool MoveNext()
        {
            if (_enumerator.MoveNext())
            {
                _current = _enumerator.Current;
                return true;
            }
            return false;
        }

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

        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        // IEnumerable<Customer> implementation
        public IEnumerator<Customer> GetEnumerator() => _source.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    class Program
    {
        // Example of a lazy enumerable that could stream millions of records.
        private static IEnumerable<Customer> GetLargeCustomerEnumerable()
        {
            for (int i = 1; i <= 1_000_000; i++)
            {
                yield return new Customer($"Customer {i}", $"Address {i}");
            }
        }

        static void Main()
        {
            // Prepare the custom data source that streams the data.
            var customers = GetLargeCustomerEnumerable();
            var dataSource = new CustomerDataSource(customers);

            Console.WriteLine("First 10 merged customers:");
            int count = 0;
            while (dataSource.MoveNext())
            {
                dataSource.GetValue("FullName", out var fullName);
                dataSource.GetValue("Address", out var address);

                if (count < 10)
                {
                    Console.WriteLine($"{fullName}, {address}");
                }

                count++;
            }

            Console.WriteLine($"Total records processed: {count}");
        }
    }
}
