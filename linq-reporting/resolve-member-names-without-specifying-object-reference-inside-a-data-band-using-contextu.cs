using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model.
    public class Order
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Customer
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // Prepare sample data.
        var order = new Order
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice", Age = 30 },
                new Customer { Name = "Bob",   Age = 45 },
                new Customer { Name = "Carol", Age = 27 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Data band (foreach) that iterates over Order.Customers.
        // Inside the band we refer to member names directly (Name, Age)
        // without specifying the object reference (c.).
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name: <<[Name]>>   Age: <<[Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // The root object is 'order' and its name in the template is "order".
        // The template uses the property 'Customers' of the root object.
        engine.BuildReport(loadedDoc, order, "order");

        // Save the generated report.
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
