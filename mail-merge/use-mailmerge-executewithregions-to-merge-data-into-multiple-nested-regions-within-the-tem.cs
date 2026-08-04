using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeNestedRegionsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ------------------------------------------------------------
            // Build the mail‑merge template with nested regions.
            // ------------------------------------------------------------
            // Outer region – Customers.
            builder.InsertField(" MERGEFIELD TableStart:Customers");

            // Fields inside the Customers region.
            builder.Write("Full name:\t");
            builder.InsertField(" MERGEFIELD FullName ");
            builder.Write("\nAddress:\t");
            builder.InsertField(" MERGEFIELD Address ");
            builder.Write("\nOrders:\n");

            // Inner region – Orders (nested inside Customers).
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Write("\tItem name:\t");
            builder.InsertField(" MERGEFIELD Name ");
            builder.Write("\n\tQuantity:\t");
            builder.InsertField(" MERGEFIELD Quantity ");
            builder.InsertParagraph();

            // Close the Orders region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders");
            // Close the Customers region.
            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // ------------------------------------------------------------
            // Prepare hierarchical data: customers each having a list of orders.
            // ------------------------------------------------------------
            CustomerList customers = new CustomerList();
            customers.Add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.Add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

            // Orders for the first customer.
            customers[0].Orders.Add(new Order("Rugby World Cup Cap", 2));
            customers[0].Orders.Add(new Order("Rugby World Cup Ball", 1));
            // Orders for the second customer.
            customers[1].Orders.Add(new Order("Rugby World Cup Guide", 1));

            // Wrap the data in a custom mail‑merge data source.
            CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

            // ------------------------------------------------------------
            // Perform the mail merge with nested regions.
            // ------------------------------------------------------------
            doc.MailMerge.ExecuteWithRegions(customersDataSource);

            // ------------------------------------------------------------
            // Save the result.
            // ------------------------------------------------------------
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "NestedMailMergeResult.docx");
            doc.Save(outputPath);

            // Indicate completion (no user interaction required).
            Console.WriteLine("Mail merge completed. Document saved to:");
            Console.WriteLine(outputPath);
        }
    }

    // ------------------------------------------------------------------------
    // Data entity classes.
    // ------------------------------------------------------------------------
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
        public List<Order> Orders { get; set; }
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

    // ------------------------------------------------------------------------
    // Typed collection for customers (required by the custom data source).
    // ------------------------------------------------------------------------
    public class CustomerList : ArrayList
    {
        public new Customer this[int index]
        {
            get { return (Customer)base[index]; }
            set { base[index] = value; }
        }
    }

    // ------------------------------------------------------------------------
    // Custom mail‑merge data source for the outer Customers region.
    // ------------------------------------------------------------------------
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly CustomerList _customers;
        private int _recordIndex = -1;

        public CustomerMailMergeDataSource(CustomerList customers)
        {
            _customers = customers;
        }

        // The name of the data source – must match the region name.
        public string TableName => "Customers";

        // Provide values for fields in the Customers region.
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

        // Move to the next customer record.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;
            return !IsEof;
        }

        // Return a child data source when the mail merge engine encounters the nested Orders region.
        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            if (tableName == "Orders")
                return new OrderMailMergeDataSource(_customers[_recordIndex].Orders);
            return null;
        }

        private bool IsEof => _recordIndex >= _customers.Count;
    }

    // ------------------------------------------------------------------------
    // Custom mail‑merge data source for the inner Orders region.
    // ------------------------------------------------------------------------
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

        public IMailMergeDataSource GetChildDataSource(string tableName) => null;

        private bool IsEof => _recordIndex >= _orders.Count;
    }
}
