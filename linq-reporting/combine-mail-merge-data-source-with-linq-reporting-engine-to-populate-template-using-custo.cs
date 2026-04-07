using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with both MailMerge fields and LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // MailMerge field – will be filled from a DataTable later.
        builder.InsertField(" MERGEFIELD CustomerName ");

        // Add a line break.
        builder.Writeln();

        // LINQ Reporting tag – will be filled from a LINQ data source later.
        builder.Writeln("Address: <<[customer.Address]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template back (required before using the ReportingEngine).
        Document doc = new Document(templatePath);

        // 3. Prepare a MailMerge data source (DataTable) and execute MailMerge.
        DataTable mailMergeTable = new DataTable("Customers");
        mailMergeTable.Columns.Add("CustomerName");
        mailMergeTable.Rows.Add("John Doe"); // Sample row.

        doc.MailMerge.Execute(mailMergeTable);

        // 4. Prepare a LINQ Reporting data source.
        Customer customer = new Customer
        {
            Address = "123 Main St, Springfield"
        };

        // 5. Use the ReportingEngine to fill the LINQ tags.
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag prefix used in the template ("customer").
        engine.BuildReport(doc, customer, "customer");

        // 6. Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model required by the LINQ Reporting engine.
public class Customer
{
    // MailMerge does not need this property, but LINQ Reporting uses it.
    public string Address { get; set; } = string.Empty;
}
