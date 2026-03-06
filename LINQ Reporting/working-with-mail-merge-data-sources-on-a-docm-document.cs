using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeDocmExample
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

    // Custom mail merge data source that implements IMailMergeDataSource.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<Customer> _customers;
        private int _recordIndex = -1; // Position before the first record.

        public CustomerMailMergeDataSource(List<Customer> customers)
        {
            _customers = customers;
        }

        // Name of the data source (used only with mail merge regions).
        public string TableName => "Customer";

        // Move to the next record.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;

            return !IsEof;
        }

        // Return a child data source for nested regions (none in this example).
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        // Provide a value for the requested field name.
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

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCM file that contains MERGEFIELDs.
            const string sourceDocPath = @"C:\Docs\Template.docm";

            // Load the DOCM document.
            Document doc = new Document(sourceDocPath);

            // Prepare sample data.
            var customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London"),
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
            };

            // Wrap the list in a custom data source.
            var dataSource = new CustomerMailMergeDataSource(customers);

            // Execute mail merge using the custom data source.
            doc.MailMerge.Execute(dataSource);

            // Optionally, adjust mail merge settings so that Word knows the document
            // is a mail‑merge main document (useful when the file is opened in Word).
            doc.MailMergeSettings.MainDocumentType = Aspose.Words.Settings.MailMergeMainDocumentType.FormLetters;
            doc.MailMergeSettings.DataType = Aspose.Words.Settings.MailMergeDataType.None; // No external data source.

            // Save the merged document. DOCX is used for the output, but you can also save as DOCM.
            const string outputPath = @"C:\Docs\MergedResult.docx";
            doc.Save(outputPath);
        }
    }
}
