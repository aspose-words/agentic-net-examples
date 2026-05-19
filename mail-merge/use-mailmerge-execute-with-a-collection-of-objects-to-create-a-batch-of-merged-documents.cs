using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add merge fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField(" MERGEFIELD FullName ");
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD Address ");

        // Prepare a collection of data objects.
        List<Customer> customers = new List<Customer>
        {
            new Customer("Thomas Hardy", "120 Hanover Sq., London"),
            new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
        };

        // Wrap the collection in a custom mail‑merge data source.
        CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

        // Execute the mail merge for all records in the collection.
        doc.MailMerge.Execute(dataSource);

        // Save the merged document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedDocument.docx");
        doc.Save(outputPath);
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

// Custom mail‑merge data source that iterates over a collection of Customer objects.
public class CustomerMailMergeDataSource : IMailMergeDataSource
{
    private readonly IList<Customer> _customers;
    private int _recordIndex = -1; // Position before the first record.

    public CustomerMailMergeDataSource(IList<Customer> customers)
    {
        _customers = customers;
    }

    // Name of the data source (used only for mail‑merge regions).
    public string TableName => "Customer";

    // Moves to the next record; returns false when the end is reached.
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

    // No child data sources are used in this example.
    public IMailMergeDataSource GetChildDataSource(string tableName) => null;

    private bool IsEof => _recordIndex >= _customers.Count;
}
