using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AsposeWordsNestedForeachExample
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
        public List<Order> Orders { get; set; }
    }

    // Simple data entity representing an order belonging to a customer.
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

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample data – a list of customers each with a list of orders.
            // -----------------------------------------------------------------
            List<Customer> customers = new List<Customer>
            {
                new Customer("Thomas Hardy", "120 Hanover Sq., London")
                {
                    Orders =
                    {
                        new Order("Rugby World Cup Cap", 2),
                        new Order("Rugby World Cup Ball", 1)
                    }
                },
                new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino")
                {
                    Orders =
                    {
                        new Order("Rugby World Cup Guide", 1)
                    }
                }
            };

            // -----------------------------------------------------------------
            // 2. Create a new blank document and a DocumentBuilder to write into it.
            // -----------------------------------------------------------------
            Document doc = new Document();                     // create document
            DocumentBuilder builder = new DocumentBuilder(doc); // create builder

            // -----------------------------------------------------------------
            // 3. Iterate over the customers (outer foreach) and write a heading.
            //    Inside that loop, iterate over each customer's orders (inner foreach)
            //    and write the order details.
            // -----------------------------------------------------------------
            foreach (Customer customer in customers)
            {
                // Write a heading for the current customer.
                builder.Writeln($"Customer: {customer.FullName}");
                builder.Writeln($"Address: {customer.Address}");
                builder.Writeln("Orders:");

                // Inner loop – list all orders for this customer.
                foreach (Order order in customer.Orders)
                {
                    builder.Writeln($"\tItem: {order.Name}, Quantity: {order.Quantity}");
                }

                // Add an empty line between customers for readability.
                builder.Writeln();
            }

            // -----------------------------------------------------------------
            // 4. Save the document to disk.
            // -----------------------------------------------------------------
            doc.Save("NestedForeachResult.docx"); // save document
        }
    }
}
