using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create the template document.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Template with a foreach loop and an if‑else block.
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        builder.Writeln("<<if [p.Quantity > 0]>>Quantity: <<[p.Quantity]>> <<else>>Out of stock<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Step 2: Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Step 3: Prepare sample data.
        ReportModel model = new ReportModel();
        model.Products.Add(new Product { Name = "Apple", Quantity = 5 });
        model.Products.Add(new Product { Name = "Banana", Quantity = 0 });

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
