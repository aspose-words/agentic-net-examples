using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace NestedMailMergeExample
{
    // Entry point of the console application.
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the start of the outer mail merge region "Customers".
            builder.InsertField(" MERGEFIELD TableStart:Customers");

            // Fields inside the "Customers" region.
            builder.Write("Full name:\t");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write("\nAddress:\t");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write("\nOrders:\n");

            // Insert the start of the nested mail merge region "Orders".
            builder.InsertField(" MERGEFIELD TableStart:Orders");

            // Fields inside the "Orders" region.
            builder.Write("\tItem name:\t");
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write("\n\tQuantity:\t");
            builder.InsertField(" MERGEFIELD Quantity ");
            builder.InsertParagraph();

            // End tags for the nested and outer regions.
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // Prepare hierarchical data: customers each with a list of orders.
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // Add orders for the first customer.
            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));

            // Add an order for the second customer.
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // Wrap the data in a custom mail merge data source.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Perform the mail merge with nested regions.
            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            // Save the result to a file in the current directory.
            doc.Save("NestedMailMerge.docx");
        }
    }

    // Simple data entity representing a customer.
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

    // Simple data entity representing an order.
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

    // Collection of Customer objects that derives from ArrayList.
    public class CustomerList : ArrayList
    {
        public new Customer this[int index]
        {
            get { return (Customer)base[index]; }
            set { base[index] = value; }
        }
    }

    // Custom mail merge data source for the "Customers" region.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly CustomerList _customers;
        private int _recordIndex = -1;

        public CustomerMailMergeDataSource(CustomerList customers)
        {
            _customers = customers;
        }

        // Name of the data source (used by Aspose.Words for region matching).
        public string TableName => "Customers";

        // Provides values for fields in the "Customers" region.
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

        // Moves to the next customer record.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;

            return !IsEof;
        }

        // Returns a child data source for the nested "Orders" region.
        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            if (tableName.Equals("Orders", StringComparison.OrdinalIgnoreCase))
                return new OrderMailMergeDataSource(_customers[_recordIndex].Orders);
            return null;
        }

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    // Custom mail merge data source for the "Orders" region.
    public class OrderMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<Order> _orders;
        private int _recordIndex = -1;

        public OrderMailMergeDataSource(List<Order> orders)
        {
            _orders = orders;
        }

        public string TableName => "Orders";

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

        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;

            return !IsEof;
        }

        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            // No further nested regions.
            return null;
        }

        private bool IsEof => _recordIndex >= _orders.Count;
    }
}
