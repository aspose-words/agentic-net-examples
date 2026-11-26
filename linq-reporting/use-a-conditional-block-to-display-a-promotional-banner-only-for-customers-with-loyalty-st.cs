using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Customer
{
    public string Name { get; set; } = string.Empty;
    public bool IsLoyal { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Dear <<[customer.Name]>>,");
        builder.Writeln();
        builder.Writeln("<<if [customer.IsLoyal]>>");
        builder.Writeln("We appreciate your loyalty! Enjoy a special 20% discount on your next purchase.");
        builder.Writeln("<</if>>");
        builder.Writeln();
        builder.Writeln("Thank you for being with us.");

        templateDoc.Save(templatePath);

        // Sample data: a loyal customer and a non‑loyal customer.
        var loyalCustomer = new Customer { Name = "Alice", IsLoyal = true };
        var nonLoyalCustomer = new Customer { Name = "Bob", IsLoyal = false };

        // Build report for the loyal customer.
        var docForLoyal = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(docForLoyal, loyalCustomer, "customer");
        string loyalReportPath = Path.Combine(outputDir, "report_loyal.docx");
        docForLoyal.Save(loyalReportPath);

        // Build report for the non‑loyal customer.
        var docForNonLoyal = new Document(templatePath);
        engine.BuildReport(docForNonLoyal, nonLoyalCustomer, "customer");
        string nonLoyalReportPath = Path.Combine(outputDir, "report_nonloyal.docx");
        docForNonLoyal.Save(nonLoyalReportPath);

        // Indicate completion.
        Console.WriteLine($"Reports generated:{Environment.NewLine}{loyalReportPath}{Environment.NewLine}{nonLoyalReportPath}");
    }
}
