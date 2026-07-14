using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    // Simple data model used by the LINQ Reporting template.
    public class Order
    {
        public string Product { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
        public List<Order> Orders { get; set; } = new();
    }

    public static void Main()
    {
        // 1. Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a title and a placeholder for the customer's name.
        builder.Writeln("Customer: <<[model.Name]>>");
        builder.Writeln("Orders:");

        // Begin a foreach block that iterates over the Orders collection.
        builder.Writeln("<<foreach [order in model.Orders]>>");
        builder.Writeln("- <<[order.Product]>> x <<[order.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // 2. Prepare sample data that matches the template.
        ReportModel model = new()
        {
            Name = "John Doe",
            Orders = new()
            {
                new Order { Product = "Apple", Quantity = 5 },
                new Order { Product = "Banana", Quantity = 3 },
                new Order { Product = "Orange", Quantity = 7 }
            }
        };

        // 3. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None // No special options required.
        };
        bool success = engine.BuildReport(template, model, "model");

        // Ensure the report was built successfully.
        if (!success)
        {
            Console.WriteLine("Report generation failed.");
            return;
        }

        // 4. Write the generated document to a memory stream (simulating an HTTP response stream).
        using (MemoryStream memoryStream = new())
        {
            // Save the document to the stream in DOCX format.
            template.Save(memoryStream, SaveFormat.Docx);

            // For demonstration, show the size of the generated document.
            Console.WriteLine($"Report written to stream. Length: {memoryStream.Length} bytes.");
        }
    }
}
