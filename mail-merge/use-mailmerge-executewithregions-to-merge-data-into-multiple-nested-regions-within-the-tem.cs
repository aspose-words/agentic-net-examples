using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to construct the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Define the outer mail merge region "Customers" -----
        builder.InsertField(" MERGEFIELD TableStart:Customers");

        // Fields that belong to the "Customers" region.
        builder.Write("Full name:\t");
        builder.InsertField(" MERGEFIELD FullName ");
        builder.Write("\nAddress:\t");
        builder.InsertField(" MERGEFIELD Address ");
        builder.Write("\nOrders:\n");

        // ----- Define the inner mail merge region "Orders" (nested inside "Customers") -----
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Write("\tItem name:\t");
        builder.InsertField(" MERGEFIELD Name ");
        builder.Write("\n\tQuantity:\t");
        builder.InsertField(" MERGEFIELD Quantity ");
        builder.InsertParagraph();

        // Close the inner region.
        builder.InsertField(" MERGEFIELD TableEnd:Orders");
        // Close the outer region.
        builder.InsertField(" MERGEFIELD TableEnd:Customers");

        // ----- Prepare hierarchical data (customers with their orders) -----
        CustomerList customers = new CustomerList();
        customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        // Add orders for each customer.
        customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
        customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
        customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

        // Wrap the data in a custom mail‑merge data source.
        CustomerMailMergeDataSource dataSource = new CustomerMailMergeDataSource(customers);

        // Perform the mail merge with nested regions.
        doc.MailMerge.ExecuteWithRegions(dataSource);

        // Save the merged document.
        doc.Save("NestedMailMerge.docx");
    }
}

// ----- Data entity classes -----
public class Customer
{
    public Customer(string fullName, string address)
    {
        FullName = fullName;
        Address = address;
        Orders = new List<Order>();
    }

    public string FullName { get; set; }
    public string Address { get; set; }
    public List<Order> Orders { get; }
}

public class Order
{
    public Order(string name, int quantity)
    {
        Name = name;
        Quantity = quantity;
    }

    public string Name { get; set; }
    public int Quantity { get; set; }
}

// ----- Typed collections -----
public class CustomerList : ArrayList
{
    public new Customer this[int index]
    {
        get => (Customer)base[index];
        set => base[index] = value;
    }
}

// ----- Custom mail‑merge data source for customers (outer region) -----
public class CustomerMailMergeDataSource : IMailMergeDataSource
{
    private readonly CustomerList _customers;
    private int _recordIndex = -1;

    public CustomerMailMergeDataSource(CustomerList customers)
    {
        _customers = customers;
    }

    // The name of the data source (must match the outer region name).
    public string TableName => "Customers";

    // Move to the next customer record.
    public bool MoveNext()
    {
        if (!IsEof)
            _recordIndex++;
        return !IsEof;
    }

    // Provide values for fields in the "Customers" region.
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

    // When the engine encounters the nested "Orders" region, return a data source for it.
    public IMailMergeDataSource GetChildDataSource(string tableName)
    {
        if (tableName.Equals("Orders", StringComparison.OrdinalIgnoreCase))
        {
            return new OrderMailMergeDataSource(_customers[_recordIndex].Orders);
        }
        return null;
    }

    private bool IsEof => _recordIndex >= _customers.Count;
}

// ----- Custom mail‑merge data source for orders (inner region) -----
public class OrderMailMergeDataSource : IMailMergeDataSource
{
    private readonly List<Order> _orders;
    private int _recordIndex = -1;

    public OrderMailMergeDataSource(List<Order> orders)
    {
        _orders = orders;
    }

    // The name of the data source (must match the inner region name).
    public string TableName => "Orders";

    // Move to the next order record.
    public bool MoveNext()
    {
        if (!IsEof)
            _recordIndex++;
        return !IsEof;
    }

    // Provide values for fields in the "Orders" region.
    public bool GetValue(string fieldName, out object fieldValue)
    {
        switch (fieldName)
        {
            case "Name":
                fieldValue = _orders[_recordIndex].Name;
                return true;
            case "Quantity":
                fieldValue = _orders[_recordIndex].Quantity;
                return true;
            default:
                fieldValue = null;
                return false;
        }
    }

    // No further nested regions.
    public IMailMergeDataSource GetChildDataSource(string tableName) => null;

    private bool IsEof => _recordIndex >= _orders.Count;
}
