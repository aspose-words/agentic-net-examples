using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

// Register code page provider for potential encoding needs.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple paragraph with a mail‑merge placeholder.
        builder.Writeln("Customer: $customer$");

        // LINQ Reporting foreach block for a collection of orders.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Order ID");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Table row that will be repeated for each order.
        builder.InsertCell();
        builder.Writeln("<<[order.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Amount]>>");
        builder.EndRow();

        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        ReportModel model = new()
        {
            Customer = "Acme Corp",
            Orders = new()
            {
                new Order { Id = 1, Amount = 100.0 },
                new Order { Id = 2, Amount = 250.5 },
                new Order { Id = 3, Amount = 75.25 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Execute mail merge to replace $customer$ fields.
        // -----------------------------------------------------------------
        doc.MailMerge.Execute(new[] { "customer" }, new object[] { model.Customer });

        // -----------------------------------------------------------------
        // 4. Build the LINQ Reporting sections.
        // -----------------------------------------------------------------
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the final report.
        // -----------------------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Public data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public string Customer { get; set; } = string.Empty;
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int Id { get; set; }
    public double Amount { get; set; }
}
