using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Mail merge field (will be filled by MailMerge).
        builder.InsertField(" MERGEFIELD CustomerName ");
        builder.Writeln(); // New paragraph.

        // LINQ Reporting tags (will be processed by ReportingEngine).
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Customer (LINQ): <<[c.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and apply MailMerge.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare a simple DataTable as the mail‑merge data source.
        DataTable mailMergeTable = new DataTable("MailMerge");
        mailMergeTable.Columns.Add("CustomerName");
        mailMergeTable.Rows.Add("John Doe");

        // Execute mail merge – fills the MERGEFIELD.
        doc.MailMerge.Execute(mailMergeTable);

        // -----------------------------------------------------------------
        // 3. Prepare LINQ Reporting data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "John Doe" },
                new Customer { Name = "Jane Smith" }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model for LINQ Reporting.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Customer
{
    public string Name { get; set; } = string.Empty;
}
