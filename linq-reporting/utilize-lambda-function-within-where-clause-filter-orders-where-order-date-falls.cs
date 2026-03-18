using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsLambdaFilterExample
{
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
        public Order(string name, int quantity, DateTime orderDate)
        {
            Name = name;
            Quantity = quantity;
            OrderDate = orderDate;
        }

        public string Name { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }

    // Custom mail‑merge data source for the root "Customers" table.
    public class CustomerMailMergeDataSource : IMailMergeDataSource
    {
        private readonly IList<Customer> _customers;
        private int _recordIndex = -1;

        public CustomerMailMergeDataSource(IList<Customer> customers)
        {
            _customers = customers;
        }

        // Name of the data source – must match the mail‑merge region name.
        public string TableName => "Customers";

        // Return field values for the current customer record.
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
                case "Orders":
                    // The "Orders" field is a child data source; we return null here.
                    fieldValue = null;
                    return false;
                default:
                    fieldValue = null;
                    return false;
            }
        }

        // Move to the next customer record.
        public bool MoveNext()
        {
            if (_recordIndex < _customers.Count - 1)
                _recordIndex++;
            return _recordIndex < _customers.Count;
        }

        // Provide the child data source for the "Orders" region.
        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            if (tableName == "Orders")
            {
                // Filter orders using a lambda expression: only orders from the last month.
                IEnumerable<Order> recentOrders = _customers[_recordIndex].Orders
                    .Where(o => o.OrderDate >= DateTime.Now.AddMonths(-1));

                return new OrderMailMergeDataSource(recentOrders);
            }

            return null;
        }
    }

    // Custom mail‑merge data source for the "Orders" child region.
    public class OrderMailMergeDataSource : IMailMergeDataSource
    {
        private readonly IList<Order> _orders;
        private int _recordIndex = -1;

        public OrderMailMergeDataSource(IEnumerable<Order> orders)
        {
            // Materialize the filtered sequence to a list for stable indexing.
            _orders = orders.ToList();
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
                case "OrderDate":
                    fieldValue = _orders[_recordIndex].OrderDate.ToString("d");
                    return true;
                default:
                    fieldValue = null;
                    return false;
            }
        }

        public bool MoveNext()
        {
            if (_recordIndex < _orders.Count - 1)
                _recordIndex++;
            return _recordIndex < _orders.Count;
        }

        public IMailMergeDataSource GetChildDataSource(string tableName) => null;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 2. Build the mail‑merge template with nested regions.
            //    Outer region: Customers
            builder.InsertField(" MERGEFIELD TableStart:Customers");
            builder.Write("Customer: ");
            builder.InsertField(" MERGEFIELD FullName");
            builder.Writeln();
            builder.Write("Address: ");
            builder.InsertField(" MERGEFIELD Address");
            builder.Writeln();
            builder.Writeln("Orders placed in the last month:");

            //    Inner region: Orders (will be filtered by lambda)
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Write("\tItem: ");
            builder.InsertField(" MERGEFIELD Name");
            builder.Write("\tQty: ");
            builder.InsertField(" MERGEFIELD Quantity");
            builder.Write("\tDate: ");
            builder.InsertField(" MERGEFIELD OrderDate");
            builder.Writeln();
            builder.InsertField(" MERGEFIELD TableEnd:Orders");

            builder.InsertField(" MERGEFIELD TableEnd:Customers");

            // 3. Prepare sample data.
            var customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London")
                {
                    Orders =
                    {
                        new Order("Rugby World Cup Cap", 2, DateTime.Now.AddDays(-10)), // within last month
                        new Order("Rugby World Cup Ball", 1, DateTime.Now.AddMonths(-2)) // older than a month
                    }
                },
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
                {
                    Orders =
                    {
                        new Order("Rugby World Cup Guide", 1, DateTime.Now.AddDays(-5)) // within last month
                    }
                }
            };

            // 4. Execute mail merge with regions, using the custom data source.
            var dataSource = new CustomerMailMergeDataSource(customers);
            doc.MailMerge.ExecuteWithRegions(dataSource);

            // 5. Save the resulting document.
            doc.Save("FilteredOrders.docx");
        }
    }
}
