using System;
using System.Collections.Generic;
using System.Linq;

namespace LambdaFilterExample
{
    // Simple data entity representing a customer.
    public class Customer
    {
        public string FullName { get; set; }
        public string Address { get; set; }
        public List<Order> Orders { get; } = new List<Order>();
    }

    // Simple data entity representing an order.
    public class Order
    {
        public string Name { get; }
        public int Quantity { get; }
        public DateTime OrderDate { get; }

        public Order(string name, int quantity, DateTime orderDate)
        {
            Name = name;
            Quantity = quantity;
            OrderDate = orderDate;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            var customers = new List<Customer>
            {
                new Customer
                {
                    FullName = "Thomas Hardy",
                    Address = "120 Hanover Sq., London",
                    Orders =
                    {
                        new Order("Rugby World Cup Cap", 2, DateTime.Now.AddDays(-10)), // within last month
                        new Order("Rugby World Cup Ball", 1, DateTime.Now.AddMonths(-2)) // older than a month
                    }
                },
                new Customer
                {
                    FullName = "Paolo Accorti",
                    Address = "Via Monte Bianco 34, Torino",
                    Orders =
                    {
                        new Order("Rugby World Cup Guide", 1, DateTime.Now.AddDays(-5)) // within last month
                    }
                }
            };

            // 2. Define the cutoff date (orders from the last month).
            DateTime cutoffDate = DateTime.Now.AddMonths(-1);

            // 3. Iterate through customers and display orders that satisfy the lambda filter.
            foreach (var customer in customers)
            {
                Console.WriteLine($"Customer: {customer.FullName}");
                Console.WriteLine($"Address: {customer.Address}");

                // Lambda filter: only orders whose OrderDate is on or after the cutoff.
                var recentOrders = customer.Orders
                    .Where(o => o.OrderDate >= cutoffDate)
                    .ToList();

                if (recentOrders.Count == 0)
                {
                    Console.WriteLine("  No orders in the last month.");
                }
                else
                {
                    Console.WriteLine("  Orders placed in the last month:");
                    foreach (var order in recentOrders)
                    {
                        Console.WriteLine($"    Item: {order.Name}, Qty: {order.Quantity}, Date: {order.OrderDate:d}");
                    }
                }

                Console.WriteLine();
            }
        }
    }
}
