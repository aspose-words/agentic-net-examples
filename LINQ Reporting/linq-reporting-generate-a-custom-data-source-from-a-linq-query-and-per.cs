using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains MERGEFIELD tags: FullName and City.
        Document template = new Document("Template.dotx");

        // Sample data collection.
        List<Customer> allCustomers = new List<Customer>
        {
            new Customer("Thomas Hardy", "London"),
            new Customer("Paolo Accorti", "Turin"),
            new Customer("John Doe", "New York")
        };

        // LINQ query to obtain only customers whose city starts with 'L'.
        List<Customer> filteredCustomers = allCustomers
            .Where(c => c.City.StartsWith("L"))
            .ToList();

        // Wrap the LINQ result in a custom mail‑merge data source.
        CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(filteredCustomers);

        // Perform the mail merge using the custom data source.
        template.MailMerge.Execute(dataSource);

        // Save the merged document.
        template.Save("Result.docx");
    }
}

// Simple POCO representing a customer.
public class Customer
{
    public Customer(string fullName, string city)
    {
        FullName = fullName;
        City = city;
    }

    public string FullName { get; set; }
    public string City { get; set; }
}

// Custom data source that implements IMailMergeDataSource.
// It supplies values for the MERGEFIELD tags defined in the template.
public class CustomerMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<Customer> _customers;
    private int _recordIndex = -1; // Position before the first record.

    public CustomerMailMergeDataSource(List<Customer> customers)
    {
        _customers = customers;
    }

    // Name of the data source (used only when mail merge regions are present).
    public string TableName => "Customer";

    // Returns the value for a given field name.
    public bool GetValue(string fieldName, out object fieldValue)
    {
        switch (fieldName)
        {
            case "FullName":
                fieldValue = _customers[_recordIndex].FullName;
                return true;
            case "City":
                fieldValue = _customers[_recordIndex].City;
                return true;
            default:
                fieldValue = null;
                return false; // Field not found.
        }
    }

    // Moves to the next record in the collection.
    public bool MoveNext()
    {
        if (_recordIndex < _customers.Count - 1)
            _recordIndex++;

        return _recordIndex < _customers.Count;
    }

    // No child data sources are required for this simple example.
    public IMailMergeDataSource GetChildDataSource(string tableName) => null;
}
