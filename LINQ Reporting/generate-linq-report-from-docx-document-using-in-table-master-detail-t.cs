using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Master entity
    public class Customer
    {
        public string Name { get; set; }
        public string City { get; set; }

        // Detail collection
        public List<Order> Orders { get; set; } = new List<Order>();
    }

    // Detail entity
    public class Order
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOCX template that contains the in‑table master‑detail tags.
            //    Example tags in the template:
            //    <<foreach [customers]>>
            //        <<[Name]>> <<[City]>>
            //        <<foreach [Orders]>>
            //            <<[Product]>> <<[Quantity]>> <<[Price]:currency>>
            //        <</foreach>>
            //    <</foreach>>
            Document template = new Document("MasterDetailTemplate.docx");

            // 2. Prepare the hierarchical data source (master with nested detail collection).
            List<Customer> customers = new List<Customer>
            {
                new Customer
                {
                    Name = "John Doe",
                    City = "Seattle",
                    Orders = new List<Order>
                    {
                        new Order { Product = "Laptop", Quantity = 1, Price = 1299.99m },
                        new Order { Product = "Mouse",  Quantity = 2, Price = 25.50m }
                    }
                },
                new Customer
                {
                    Name = "Jane Smith",
                    City = "New York",
                    Orders = new List<Order>
                    {
                        new Order { Product = "Desk",   Quantity = 1, Price = 250.00m },
                        new Order { Product = "Chair",  Quantity = 4, Price = 85.75m },
                        new Order { Product = "Lamp",   Quantity = 2, Price = 45.00m }
                    }
                }
            };

            // 3. Build the report using the LINQ Reporting Engine.
            //    The data source name ("customers") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, customers, "customers");

            // 4. Save the populated document.
            template.Save("MasterDetailReport.docx");
        }
    }
}
