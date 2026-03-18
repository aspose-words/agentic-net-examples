using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Build a template document that contains two nested mail‑merge regions:
        //   - outer region  : Customers
        //   - inner region  : Orders (child of Customers)
        Document doc = CreateTemplate();

        // Prepare hierarchical data.
        CustomerList customers = new CustomerList();
        customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
        customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
        customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

        // Wrap the top‑level collection in a custom data source that implements IMailMergeDataSource.
        CustomerMailMergeDataSource rootDataSource = new CustomerMailMergeDataSource(customers);

        // Execute the nested mail merge. Aspose.Words will call GetChildDataSource for the inner region.
        doc.MailMerge.ExecuteWithRegions(rootDataSource);

        // Save the merged document.
        doc.Save("NestedMailMerge.docx");
    }

    // Creates a blank document and inserts the required TableStart/TableEnd fields.
    static Document CreateTemplate()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Customers region (outer) -----
        builder.InsertField(" MERGEFIELD TableStart:Customers");
        builder.Writeln("Full name: ");
        builder.InsertField(" MERGEFIELD FullName");
        builder.Writeln();
        builder.Writeln("Address: ");
        builder.InsertField(" MERGEFIELD Address");
        builder.Writeln("Orders:");

        // ----- Orders region (inner) -----
        builder.InsertField(" MERGEFIELD TableStart:Orders");
        builder.Write("\tItem: ");
        builder.InsertField(" MERGEFIELD Name");
        builder.Write("\tQty: ");
        builder.InsertField(" MERGEFIELD Quantity");
        builder.Writeln();
        builder.InsertField(" MERGEFIELD TableEnd:Orders");

        // ----- End of Customers region -----
        builder.InsertField(" MERGEFIELD TableEnd:Customers");

        return doc;
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

    // ----- Collection wrapper for customers (required for ArrayList inheritance) -----
    public class CustomerList : ArrayList
    {
        public new Customer this[int index]
        {
            get => (Customer)base[index];
            set => base[index] = value;
        }
    }

    // ----- Root data source (Customers) -----
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly CustomerList _customers;
        private int _recordIndex = -1; // positioned before the first record

        public CustomerMailMergeDataSource(CustomerList customers)
        {
            _customers = customers;
        }

        public string TableName => "Customers";

        public bool MoveNext()
        {
            if (!IsEof) _recordIndex++;
            return !IsEof;
        }

        private bool IsEof => _recordIndex >= _customers.Count;

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

        // When the engine encounters TableStart:Orders, it asks for a child data source.
        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            if (string.Equals(tableName, "Orders", StringComparison.OrdinalIgnoreCase))
                return new OrderMailMergeDataSource(_customers[_recordIndex].Orders);
            return null;
        }
    }

    // ----- Child data source (Orders) -----
    public class OrderMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<Order> _orders;
        private int _recordIndex = -1; // positioned before the first record

        public OrderMailMergeDataSource(List<Order> orders)
        {
            _orders = orders;
        }

        public string TableName => "Orders";

        public bool MoveNext()
        {
            if (!IsEof) _recordIndex++;
            return !IsEof;
        }

        private bool IsEof => _recordIndex >= _orders.Count;

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

        // Orders have no further nested regions.
        public IMailMergeDataSource GetChildDataSource(string tableName) => null;
    }
}
