using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExternalTypeDemo
{
    // Business object that we want to expose directly in the mail‑merge template.
    public class CustomerInfo
    {
        public CustomerInfo(string fullName, string address, string phone)
        {
            FullName = fullName;
            Address  = address;
            Phone    = phone;
        }

        public string FullName { get; set; }
        public string Address  { get; set; }
        public string Phone    { get; set; }
    }

    // Custom mail‑merge data source that wraps a collection of CustomerInfo objects.
    // Implements IMailMergeDataSource so Aspose.Words can query field values by name.
    public class CustomerInfoMailMergeDataSource : IMailMergeDataSource
    {
        private readonly IList<CustomerInfo> _customers;
        private int _recordIndex = -1; // Position before the first record.

        public CustomerInfoMailMergeDataSource(IList<CustomerInfo> customers)
        {
            _customers = customers ?? throw new ArgumentNullException(nameof(customers));
        }

        // The name of the data source – used only when mail‑merge regions are involved.
        public string TableName => "CustomerInfo";

        // Returns the value of the requested field for the current record.
        public bool GetValue(string fieldName, out object fieldValue)
        {
            // Ensure we are positioned on a valid record.
            if (_recordIndex < 0 || _recordIndex >= _customers.Count)
            {
                fieldValue = null;
                return false;
            }

            var current = _customers[_recordIndex];
            switch (fieldName)
            {
                case "FullName":
                    fieldValue = current.FullName;
                    return true;
                case "Address":
                    fieldValue = current.Address;
                    return true;
                case "Phone":
                    fieldValue = current.Phone;
                    return true;
                default:
                    fieldValue = null;
                    return false; // Field not found.
            }
        }

        // Moves to the next record; returns false when the end of the collection is reached.
        public bool MoveNext()
        {
            if (_recordIndex < _customers.Count - 1)
            {
                _recordIndex++;
                return true;
            }
            return false;
        }

        // No child data sources are required for this simple example.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Build a simple mail‑merge template that references the properties of CustomerInfo.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Customer Report");
            builder.Writeln("----------------");
            builder.Writeln("Name   : ");
            builder.InsertField("MERGEFIELD FullName");   // Will be replaced by CustomerInfo.FullName
            builder.Writeln();
            builder.Writeln("Address: ");
            builder.InsertField("MERGEFIELD Address");    // Will be replaced by CustomerInfo.Address
            builder.Writeln();
            builder.Writeln("Phone  : ");
            builder.InsertField("MERGEFIELD Phone");      // Will be replaced by CustomerInfo.Phone
            builder.Writeln();

            // 3. Prepare a list of CustomerInfo objects to be merged.
            var customers = new List<CustomerInfo>
            {
                new CustomerInfo("Alice Johnson", "123 Maple St., Springfield", "555‑0101"),
                new CustomerInfo("Bob Smith",     "456 Oak Ave., Shelbyville", "555‑0202")
            };

            // 4. Wrap the list in our custom data source.
            IMailMergeDataSource dataSource = new CustomerInfoMailMergeDataSource(customers);

            // 5. Execute the mail merge. Aspose.Words will call GetValue for each field name.
            doc.MailMerge.Execute(dataSource);

            // 6. Save the merged document.
            doc.Save("CustomerInfoMailMergeResult.docx");
        }
    }
}
