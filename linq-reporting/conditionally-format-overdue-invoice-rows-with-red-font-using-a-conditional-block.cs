using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for a single invoice.
    public class Invoice
    {
        public int Id { get; set; } = 0;
        public DateTime Date { get; set; } = DateTime.Now;
        public decimal Amount { get; set; } = 0m;
        public DateTime DueDate { get; set; } = DateTime.Now;

        // Computed property indicating whether the invoice is overdue.
        public bool IsOverdue => DateTime.Now > DueDate;
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Invoice> Invoices { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            const string templateFile = "InvoiceTemplate.docx";

            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin the foreach block over the Invoices collection.
            builder.Writeln("<<foreach [inv in Invoices]>>");

            // Start a table that will be repeated for each invoice.
            Table table = builder.StartTable();

            // Table header.
            builder.InsertCell();
            builder.Writeln("Id");
            builder.InsertCell();
            builder.Writeln("Date");
            builder.InsertCell();
            builder.Writeln("Amount");
            builder.InsertCell();
            builder.Writeln("Due Date");
            builder.EndRow();

            // Data row – Id.
            builder.InsertCell();
            builder.Writeln("<<[inv.Id]>>");

            // Data row – Date.
            builder.InsertCell();
            builder.Writeln("<<[inv.Date]>>");

            // Data row – Amount with conditional red font if overdue.
            builder.InsertCell();
            builder.Writeln(
                "<<if [inv.IsOverdue]>>" +
                "<<textColor [\"Red\"]>><<[inv.Amount]>> <</textColor>><</if>>" +
                "<<if [!inv.IsOverdue]>>" +
                "<<[inv.Amount]>>" +
                "<</if>>");

            // Data row – Due Date.
            builder.InsertCell();
            builder.Writeln("<<[inv.DueDate]>>");

            // End of the data row.
            builder.EndRow();

            // End the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templateFile);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var invoices = new List<Invoice>
            {
                new Invoice { Id = 1, Date = DateTime.Today.AddDays(-30), Amount = 150.00m, DueDate = DateTime.Today.AddDays(-5) },
                new Invoice { Id = 2, Date = DateTime.Today.AddDays(-20), Amount = 250.00m, DueDate = DateTime.Today.AddDays(10) },
                new Invoice { Id = 3, Date = DateTime.Today.AddDays(-10), Amount = 350.00m, DueDate = DateTime.Today.AddDays(-1) }
            };

            var model = new ReportModel { Invoices = invoices };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templateFile);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the wrapper object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            const string outputFile = "InvoiceReport.docx";
            reportDoc.Save(outputFile);
        }
    }
}
