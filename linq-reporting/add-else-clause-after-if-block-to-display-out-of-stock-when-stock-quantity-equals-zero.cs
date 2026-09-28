using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public int Stock { get; set; }
}

public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data
        var model = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple", Stock = 10 },
                new Product { Name = "Banana", Stock = 0 },
                new Product { Name = "Orange", Stock = 5 }
            }
        };

        // Create template document
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("<<if [p.Stock > 0]>>In stock: <<[p.Stock]>> <</if>>");
        builder.Writeln("<<if [p.Stock == 0]>>Out of stock<</if>>");
        builder.Writeln("<</foreach>>");

        // Build report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save output
        doc.Save("ReportOutput.docx");
    }
}
