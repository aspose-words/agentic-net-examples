using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace NestedMailMergeExample
{
    // Data entity representing a customer.
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

    // Data entity representing an order.
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

    // Typed collection of customers required by the custom data source.
    public class CustomerList : ArrayList
    {
        public new Customer this[int index]
        {
            get { return (Customer)base[index]; }
            set { base[index] = value; }
        }
    }

    // Custom mail merge data source for the Customers table.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly CustomerList _customers;
        private int _recordIndex = -1;

        public CustomerMailMergeDataSource(CustomerList customers)
        {
            _customers = customers;
        }

        // Name of the data source (must match the region name in the template).
        public string TableName => "Customers";

        // Move to the next customer record.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;
            return !IsEof;
        }

        // Provide values for fields inside the Customers region.
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

        // Return a child data source for the nested Orders region.
        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            if (string.Equals(tableName, "Orders", StringComparison.OrdinalIgnoreCase))
                return new OrderMailMergeDataSource(_customers[_recordIndex].Orders);
            return null;
        }

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    // Custom mail merge data source for the Orders table.
    public class OrderMailMergeDataSource : IMailMergeDataSource
    {
        private readonly List<Order> _orders;
        private int _recordIndex = -1;

        public OrderMailMergeDataSource(List<Order> orders)
        {
            _orders = orders;
        }

        public string TableName => "Orders";

        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;
            return !IsEof;
        }

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

        private bool IsEof => _recordIndex >= _orders.Count;
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define the outer mail merge region for Customers.
            builder.InsertField(" MERGEFIELD TableStart:Customers ");

            // Fields inside the Customers region.
            builder.Write("Full name:\t");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write("\nAddress:\t");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write("\nOrders:\n");

            // Define the inner mail merge region for Orders.
            builder.InsertField(" MERGEFIELD TableStart:Orders ");
            builder.Write("\tItem name:\t");
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write("\n\tQuantity:\t");
            builder.InsertField(" MERGEFIELD Quantity ");
            builder.InsertParagraph();

            // Close the Orders region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders ");

            // Close the Customers region.
            builder.InsertField(" MERGEFIELD TableEnd:Customers ");

            // Populate hierarchical data.
            CustomerList customers = new CustomerList();
            Customer cust1 = new Customer("Thomas Hardy", "120 Hanover Sq., London");
            cust1.Orders.Add(new Order("Rugby World Cup Cap", 2));
            cust1.Orders.Add(new Order("Rugby World Cup Ball", 1));
            customers.Add(cust1);

            Customer cust2 = new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino");
            cust2.Orders.Add(new Order("Rugby World Cup Guide", 1));
            customers.Add(cust2);

            // Wrap the data in a custom mail merge data source.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // Execute the nested mail merge using the custom data source.
            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            // Save the resulting document.
            doc.Save("NestedMailMerge.docx");
        }
    }
}
