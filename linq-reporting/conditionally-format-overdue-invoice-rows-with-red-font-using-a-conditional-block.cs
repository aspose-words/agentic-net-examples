using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace InvoiceReportExample
{
    // Data model for a single invoice.
    public class Invoice
    {
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public decimal Amount { get; set; }
        public DateTime DueDate { get; set; }

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
            // Paths for the template and the generated report.
            string templatePath = "InvoiceTemplate.docx";
            string reportPath = "InvoiceReport.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("Invoice Report");
            builder.Writeln();

            // Begin the foreach block over Invoices.
            builder.Writeln("<<foreach [inv in Invoices]>>");

            // Start a table for each invoice (header + data row).
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("ID");
            builder.InsertCell();
            builder.Writeln("Date");
            builder.InsertCell();
            builder.Writeln("Amount");
            builder.InsertCell();
            builder.Writeln("Due Date");
            builder.EndRow();

            // Data row – each cell contains the appropriate tag.
            builder.InsertCell();
            builder.Writeln("<<[inv.Id]>>");

            builder.InsertCell();
            builder.Writeln("<<[inv.Date]>>");

            builder.InsertCell();
            // Conditional formatting: red text if overdue.
            builder.Writeln(
                "<<if [inv.IsOverdue]>> <<textColor [\"Red\"]>><<[inv.Amount]>> <</textColor>><</if>>" +
                "<<if [inv.IsOverdue == false]>> <<[inv.Amount]>> <</if>>");

            builder.InsertCell();
            builder.Writeln("<<[inv.DueDate]>>");

            // End of the data row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new()
            {
                Invoices = new()
                {
                    new Invoice
                    {
                        Id = 1,
                        Date = DateTime.Today.AddDays(-30),
                        Amount = 1500.00m,
                        DueDate = DateTime.Today.AddDays(-5) // overdue
                    },
                    new Invoice
                    {
                        Id = 2,
                        Date = DateTime.Today.AddDays(-10),
                        Amount = 750.50m,
                        DueDate = DateTime.Today.AddDays(10) // not overdue
                    },
                    new Invoice
                    {
                        Id = 3,
                        Date = DateTime.Today.AddDays(-20),
                        Amount = 320.75m,
                        DueDate = DateTime.Today.AddDays(-1) // overdue
                    }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
