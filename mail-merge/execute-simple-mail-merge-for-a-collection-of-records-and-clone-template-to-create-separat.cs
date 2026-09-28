using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a template document with merge fields.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("Dear <<FirstName>> <<LastName>>,");
            builder.Writeln("Your order <<OrderId>> is confirmed.");
            builder.Writeln("Thank you for shopping with us.");
            builder.Writeln(); // Add an empty line.

            // Define a collection of records to merge.
            var customers = new List<Customer>
            {
                new Customer { FirstName = "John",  LastName = "Doe",   OrderId = 1001 },
                new Customer { FirstName = "Jane",  LastName = "Smith", OrderId = 1002 },
                new Customer { FirstName = "Alice", LastName = "Brown", OrderId = 1003 }
            };

            // Perform mail merge for each record, cloning the template each time.
            foreach (var customer in customers)
            {
                // Clone the template to keep it unchanged for the next iteration.
                Document doc = (Document)template.Clone();

                // Execute mail merge with the current record's data.
                doc.MailMerge.Execute(
                    new[] { "FirstName", "LastName", "OrderId" },
                    new object[] { customer.FirstName, customer.LastName, customer.OrderId });

                // Save the merged document to a separate file.
                string fileName = $"MergedDocument_{customer.OrderId}.docx";
                doc.Save(fileName);
            }
        }

        // Simple data class representing a record for mail merge.
        private class Customer
        {
            public string FirstName { get; set; }
            public string LastName  { get; set; }
            public int    OrderId   { get; set; }
        }
    }
}
