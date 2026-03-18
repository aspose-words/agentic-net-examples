using System;
using System.Collections.Generic;
using System.Linq;

namespace OrderFilteringExample
{
    // Simple POCO representing an order.
    public class Order
    {
        public int Id { get; set; }
        public decimal TotalAmount { get; set; }

        public override string ToString()
        {
            return $"Order Id: {Id}, TotalAmount: {TotalAmount:C}";
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample data: a collection of orders.
            List<Order> orders = new List<Order>
            {
                new Order { Id = 1, TotalAmount = 150.00m },
                new Order { Id = 2, TotalAmount = 75.50m },
                new Order { Id = 3, TotalAmount = 250.00m },
                new Order { Id = 4, TotalAmount = 40.00m }
            };

            // Define the threshold for filtering.
            decimal threshold = 100.00m;

            // Lambda expression used inside LINQ's Where to filter orders
            // where the TotalAmount exceeds the specified threshold.
            IEnumerable<Order> filteredOrders = orders.Where(o => o.TotalAmount > threshold);

            // Output the filtered orders.
            Console.WriteLine($"Orders with TotalAmount > {threshold:C}:");
            foreach (Order order in filteredOrders)
            {
                Console.WriteLine(order);
            }
        }
    }
}
