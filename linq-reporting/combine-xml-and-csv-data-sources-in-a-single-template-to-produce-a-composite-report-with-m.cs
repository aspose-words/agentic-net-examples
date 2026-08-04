using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // ReportingEngine, XmlDataSource, CsvDataSource, CsvDataLoadOptions

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample XML data.
        // -----------------------------------------------------------------
        string xmlPath = "Products.xml";
        File.WriteAllText(xmlPath,
            @"<Products>
                <Product>
                    <Name>Apple</Name>
                    <Price>1.20</Price>
                </Product>
                <Product>
                    <Name>Banana</Name>
                    <Price>0.80</Price>
                </Product>
                <Product>
                    <Name>Cherry</Name>
                    <Price>2.50</Price>
                </Product>
              </Products>");

        // -----------------------------------------------------------------
        // 2. Create sample CSV data.
        // -----------------------------------------------------------------
        string csvPath = "Customers.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Email\r\n" +
            "1,John Doe,john.doe@example.com\r\n" +
            "2,Jane Smith,jane.smith@example.com\r\n" +
            "3,Bob Johnson,bob.johnson@example.com\r\n");

        // -----------------------------------------------------------------
        // 3. Build a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();                 // create blank document
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Composite Report");
        builder.Writeln("-----------------");
        builder.Writeln();

        // XML section – list of products.
        builder.Writeln("Products (from XML):");
        builder.Writeln("<<foreach [product in xml]>>");
        builder.Writeln("Name: <<[product.Name]>>");
        builder.Writeln("Price: $<<[product.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // CSV section – list of customers.
        builder.Writeln("Customers (from CSV):");
        builder.Writeln("<<foreach [cust in csv]>>");
        builder.Writeln("Id: <<[cust.Id]>>");
        builder.Writeln("Name: <<[cust.Name]>>");
        builder.Writeln("Email: <<[cust.Email]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (lifecycle step: create → save).
        string templatePath = "CompositeTemplate.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template back (demonstrates load step).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Prepare data source objects.
        // -----------------------------------------------------------------
        XmlDataSource xmlData = new XmlDataSource(xmlPath);
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions(true); // first row contains headers
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // -----------------------------------------------------------------
        // 6. Build the report using both data sources.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // BuildReport overload that accepts multiple sources.
        engine.BuildReport(doc, new object[] { xmlData, csvData }, new string[] { "xml", "csv" });

        // -----------------------------------------------------------------
        // 7. Save the final report.
        // -----------------------------------------------------------------
        string outputPath = "CompositeReport.docx";
        doc.Save(outputPath);

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(xmlPath);
        // File.Delete(csvPath);
        // File.Delete(templatePath);
    }
}
