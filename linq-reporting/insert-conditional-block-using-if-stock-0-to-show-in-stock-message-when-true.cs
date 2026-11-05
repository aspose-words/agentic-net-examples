using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare paths.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // Create the template document with a conditional block.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Product Stock Status:");
        builder.Writeln("<<if [product.Stock > 0]>>In stock<</if>>");
        builder.Writeln("<<if [product.Stock <= 0]>>Out of stock<</if>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        Product product = new Product { Stock = 5 };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, product, "product");

        // Save the generated report.
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Public data model class.
public class Product
{
    public int Stock { get; set; } = 0;
}
