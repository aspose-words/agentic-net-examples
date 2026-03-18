using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create an empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting template that iterates over a collection named "customers"
        // and outputs each customer's Name in uppercase using the ":upper" format switch.
        builder.Writeln("<<foreach [in customers]>>");
        builder.Writeln("<<[Name]:upper>>");
        builder.Writeln("<</foreach>>");

        // Prepare the data source – a list of Customer objects.
        List<Customer> customers = new List<Customer>
        {
            new Customer { Name = "Alice" },
            new Customer { Name = "Bob" },
            new Customer { Name = "Charlie" }
        };

        // Build the report. The second parameter is the data source,
        // the third parameter is the name used inside the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, customers, "customers");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
