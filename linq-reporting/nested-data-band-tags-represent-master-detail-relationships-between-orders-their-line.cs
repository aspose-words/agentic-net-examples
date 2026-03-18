using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace NestedMailMergeExample
{
    // Represents a single order.
    public class Order
    {
        public Order(int id, string customerName, OrderDetailList details)
        {
            OrderId = id;
            CustomerName = customerName;
            Details = details;
        }

        public int OrderId { get; }
        public string CustomerName { get; }
        public OrderDetailList Details { get; }
    }

    // Represents a line item of an order.
    public class OrderDetail
    {
        public OrderDetail(string product, int quantity)
        {
            Product = product;
            Quantity = quantity;
        }

        public string Product { get; }
        public int Quantity { get; }
    }

    // Typed collection of Order objects.
    public class OrderList : ArrayList
    {
        public new Order this[int index]
        {
            get => (Order)base[index];
            set => base[index] = value;
        }
    }

    // Typed collection of OrderDetail objects.
    public class OrderDetailList : List<OrderDetail>
    {
    }

    // Mail merge data source for the master table (Orders).
    public class OrderMailMergeDataSource : IMailMergeDataSource
    {
        private readonly OrderList _orders;
        private int _recordIndex = -1; // Position before the first record.

        public OrderMailMergeDataSource(OrderList orders)
        {
            _orders = orders;
        }

        // Name of the data source – must match the mail‑merge region name.
        public string TableName => "Orders";

        // Move to the next order record.
        public bool MoveNext()
        {
            if (!IsEof)
                _recordIndex++;
            return !IsEof;
        }

        // Provide field values for the current order.
        public bool GetValue(string fieldName, out object fieldValue)
        {
            switch (fieldName)
            {
                case "OrderId":
                    fieldValue = _orders[_recordIndex].OrderId;
                    return true;
                case "CustomerName":
                    fieldValue = _orders[_recordIndex].CustomerName;
                    return true;
                default:
                    fieldValue = null!; // null‑forgiving – never used by the engine for unknown fields.
                    return false;
            }
        }

        // Return a child data source (OrderDetails) when the engine encounters the nested region.
        public IMailMergeDataSource? GetChildDataSource(string tableName)
        {
            if (string.Equals(tableName, "OrderDetails", StringComparison.OrdinalIgnoreCase))
            {
                // Create a data source for the details of the current order.
                return new OrderDetailMailMergeDataSource(_orders[_recordIndex].Details);
            }
            // No relation – return null.
            return null;
        }

        private bool IsEof => _recordIndex >= _orders.Count;
    }

    // Mail merge data source for the detail table (OrderDetails).
    public class OrderDetailMailMergeDataSource : IMailMergeDataSource
    {
        private readonly OrderDetailList _details;
        private int _recordIndex = -1; // Position before the first record.

        public OrderDetailMailMergeDataSource(OrderDetailList details)
        {
            _details = details;
        }

        public string TableName => "OrderDetails";

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
                case "Product":
                    fieldValue = _details[_recordIndex].Product;
                    return true;
                case "Quantity":
                    fieldValue = _details[_recordIndex].Quantity;
                    return true;
                default:
                    fieldValue = null!;
                    return false;
            }
        }

        // OrderDetails have no further child tables.
        public IMailMergeDataSource? GetChildDataSource(string tableName) => null;

        private bool IsEof => _recordIndex >= _details.Count;
    }

    class Program
    {
        static void Main()
        {
            // Build a simple template document with nested mail‑merge regions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Outer region – Orders.
            builder.InsertField(" MERGEFIELD TableStart:Orders");
            builder.Write("Order ID: ");
            builder.InsertField(" MERGEFIELD OrderId");
            builder.Write("\nCustomer: ");
            builder.InsertField(" MERGEFIELD CustomerName");
            builder.Write("\nItems:\n");

            // Inner region – OrderDetails.
            builder.InsertField(" MERGEFIELD TableStart:OrderDetails");
            builder.Write("\tProduct: ");
            builder.InsertField(" MERGEFIELD Product");
            builder.Write(", Quantity: ");
            builder.InsertField(" MERGEFIELD Quantity");
            builder.InsertParagraph();
            builder.InsertField(" MERGEFIELD TableEnd:OrderDetails");

            // End of outer region.
            builder.InsertField(" MERGEFIELD TableEnd:Orders");

            // Prepare sample data.
            OrderList orders = new OrderList();

            orders.Add(new Order(
                id: 1001,
                customerName: "Alice Johnson",
                details: new OrderDetailList
                {
                    new OrderDetail("Laptop", 1),
                    new OrderDetail("Mouse", 2)
                }));

            orders.Add(new Order(
                id: 1002,
                customerName: "Bob Smith",
                details: new OrderDetailList
                {
                    new OrderDetail("Desk", 1),
                    new OrderDetail("Chair", 4),
                    new OrderDetail("Lamp", 2)
                }));

            // Wrap the collection in a mail‑merge data source.
            IMailMergeDataSource dataSource = new OrderMailMergeDataSource(orders);

            // Execute the mail merge with regions.
            doc.MailMerge.ExecuteWithRegions(dataSource);

            // Save the result.
            doc.Save("NestedMailMergeResult.docx");
        }
    }
}
