using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a template document with a master (Customers) region
        // and a nested detail (Invoices) region.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin master region.
        builder.InsertField(" MERGEFIELD TableStart:Customers");
        builder.Write("Customer: ");
        builder.InsertField(" MERGEFIELD FullName");
        builder.Writeln();
        builder.Write("Address: ");
        builder.InsertField(" MERGEFIELD Address");
        builder.Writeln();

        // Begin nested detail region.
        builder.InsertField(" MERGEFIELD TableStart:Invoices");
        builder.Write("\tInvoice #: ");
        builder.InsertField(" MERGEFIELD InvoiceNumber");
        builder.Write(", Amount: ");
        builder.InsertField(" MERGEFIELD Amount");
        builder.Writeln();
        // End nested detail region.
        builder.InsertField(" MERGEFIELD TableEnd:Invoices");

        // End master region.
        builder.InsertField(" MERGEFIELD TableEnd:Customers");

        // Prepare hierarchical data.
        CustomerList customers = new CustomerList();

        Customer cust1 = new Customer("John Doe", "123 Main St");
        cust1.Invoices.Add(new Invoice("INV001", 250.00));
        cust1.Invoices.Add(new Invoice("INV002", 125.50));
        customers.Add(cust1);

        Customer cust2 = new Customer("Jane Smith", "456 Oak Ave");
        cust2.Invoices.Add(new Invoice("INV003", 300.00));
        customers.Add(cust2);

        // Wrap the master collection in a mail‑merge data source.
        CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

        // Execute mail merge with regions – the engine will call GetChildDataSource
        // for each customer to obtain the corresponding invoices.
        doc.MailMerge.ExecuteWithRegions(dataSource);

        // Save the generated report.
        doc.Save("MasterDetailReport.docx");
    }
}

// -----------------------------------------------------------------------------
// Data entity classes.
// -----------------------------------------------------------------------------
public class Customer
{
    public Customer(string fullName, string address)
    {
        FullName = fullName;
        Address = address;
        Invoices = new List<Invoice>();
    }

    public string FullName { get; set; }
    public string Address { get; set; }
    public List<Invoice> Invoices { get; }
}

public class Invoice
{
    public Invoice(string number, double amount)
    {
        InvoiceNumber = number;
        Amount = amount;
    }

    public string InvoiceNumber { get; set; }
    public double Amount { get; set; }
}

// -----------------------------------------------------------------------------
// Typed collection for customers (required for the custom data source).
// -----------------------------------------------------------------------------
public class CustomerList : ArrayList
{
    public new Customer this[int index]
    {
        get { return (Customer)base[index]; }
        set { base[index] = value; }
    }
}

// -----------------------------------------------------------------------------
// Master data source – implements IMailMergeDataSource.
// Returns a child data source for the "Invoices" region.
// -----------------------------------------------------------------------------
public class CustomerMailMergeDataSource : IMailMergeDataSource
{
    private readonly CustomerList _customers;
    private int _recordIndex = -1; // Position before the first record.

    public CustomerMailMergeDataSource(CustomerList customers)
    {
        _customers = customers;
    }

    public string TableName => "Customers";

    public bool MoveNext()
    {
        if (_recordIndex < _customers.Count - 1)
        {
            _recordIndex++;
            return true;
        }
        return false;
    }

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

    public IMailMergeDataSource GetChildDataSource(string tableName)
    {
        // When the engine encounters TableStart:Invoices, return the detail source.
        if (tableName.Equals("Invoices", StringComparison.OrdinalIgnoreCase))
        {
            return new InvoiceMailMergeDataSource(_customers[_recordIndex].Invoices);
        }
        // No other child tables.
        return null;
    }
}

// -----------------------------------------------------------------------------
// Detail data source – implements IMailMergeDataSource.
// No further nesting, so GetChildDataSource returns null.
// -----------------------------------------------------------------------------
public class InvoiceMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<Invoice> _invoices;
    private int _recordIndex = -1; // Position before the first record.

    public InvoiceMailMergeDataSource(List<Invoice> invoices)
    {
        _invoices = invoices;
    }

    public string TableName => "Invoices";

    public bool MoveNext()
    {
        if (_recordIndex < _invoices.Count - 1)
        {
            _recordIndex++;
            return true;
        }
        return false;
    }

    public bool GetValue(string fieldName, out object fieldValue)
    {
        switch (fieldName)
        {
            case "InvoiceNumber":
                fieldValue = _invoices[_recordIndex].InvoiceNumber;
                return true;
            case "Amount":
                fieldValue = _invoices[_recordIndex].Amount;
                return true;
            default:
                fieldValue = null;
                return false;
        }
    }

    public IMailMergeDataSource GetChildDataSource(string tableName)
    {
        // No nested tables under Invoices.
        return null;
    }
}
