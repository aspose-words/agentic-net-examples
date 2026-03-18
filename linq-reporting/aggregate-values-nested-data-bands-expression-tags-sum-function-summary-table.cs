using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // <-- added namespace for Table

namespace AsposeWordsNestedAggregation
{
    // Simple data entity representing a customer.
    public class Customer
    {
        public Customer(string name)
        {
            Name = name;
            Orders = new List<Order>();
        }

        public string Name { get; set; }
        public List<Order> Orders { get; set; }
    }

    // Simple data entity representing an order line.
    public class Order
    {
        public Order(string item, double quantity)
        {
            Item = item;
            Quantity = quantity;
        }

        public string Item { get; set; }
        public double Quantity { get; set; }
    }

    // Typed collection used by the reporting engine.
    public class CustomerList : ArrayList
    {
        public new Customer this[int index]
        {
            get => (Customer)base[index];
            set => base[index] = value;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document that contains nested mail‑merge regions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table that will hold the customer and order data.
            Table table = builder.StartTable(); // Table type now resolved via Aspose.Words.Tables

            // Header row.
            builder.InsertCell();
            builder.Write("Customer");
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // ----- Begin outer region: Customers -----
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD TableStart:Customers ");
            builder.InsertField(" MERGEFIELD Name "); // Customer name.
            builder.InsertCell(); // Empty cell for inner region start.

            // ----- Begin inner region: Orders -----
            builder.InsertField(" MERGEFIELD TableStart:Orders ");
            builder.InsertField(" MERGEFIELD Item ");   // Order item.
            builder.InsertCell();
            builder.InsertField(" MERGEFIELD Quantity "); // Order quantity.
            builder.InsertField(" MERGEFIELD TableEnd:Orders ");
            // End of inner region.

            // Row that will display the sum of quantities for the current customer.
            builder.InsertCell();
            builder.Write("Total:");
            builder.InsertCell();
            // Expression tag that the ReportingEngine will evaluate: sum of Quantity in the current Orders band.
            builder.Write("<<[Orders.Quantity]:sum>>");
            builder.EndRow();

            // ----- End outer region: Customers -----
            builder.InsertField(" MERGEFIELD TableEnd:Customers ");

            builder.EndTable();

            // 2. Build the hierarchical data source.
            CustomerList customers = new CustomerList();

            Customer cust1 = new Customer("John Doe");
            cust1.Orders.Add(new Order("Widget A", 3));
            cust1.Orders.Add(new Order("Widget B", 5));
            customers.Add(cust1);

            Customer cust2 = new Customer("Jane Smith");
            cust2.Orders.Add(new Order("Gadget X", 2));
            cust2.Orders.Add(new Order("Gadget Y", 4));
            cust2.Orders.Add(new Order("Gadget Z", 1));
            customers.Add(cust2);

            // 3. Use ReportingEngine to populate the template.
            ReportingEngine engine = new ReportingEngine();
            // The data source name must match the outer region name ("Customers").
            engine.BuildReport(doc, customers, "Customers");

            // 4. Save the resulting document.
            doc.Save("AggregatedNestedReport.docx");
        }
    }
}
