using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a title.
        builder.Writeln("Customer List:");
        // Begin a foreach loop over the data source named \"Customers\".
        builder.Writeln("<<foreach [c in Customers]>>");
        // Output each customer's name and address.
        builder.Writeln("Name: <<[c.Name]>>");
        builder.Writeln("Address: <<[c.Address]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template back (required before building the report).
        Document doc = new Document(templatePath);

        // Create a DataTable that will serve as the mail‑merge‑style data source.
        DataTable customersTable = new DataTable("Customers");
        customersTable.Columns.Add("Name", typeof(string));
        customersTable.Columns.Add("Address", typeof(string));

        // Populate the table with sample data.
        customersTable.Rows.Add("Alice Johnson", "123 Maple Street");
        customersTable.Rows.Add("Bob Smith", "456 Oak Avenue");
        customersTable.Rows.Add("Carol White", "789 Pine Road");

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name must match the name used in the template tags.
        engine.BuildReport(doc, customersTable, "Customers");

        // Save the generated report.
        const string reportPath = "report.docx";
        doc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {reportPath}");
    }
}
