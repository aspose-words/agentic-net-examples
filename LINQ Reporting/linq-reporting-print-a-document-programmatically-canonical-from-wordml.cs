using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingPrint
{
    // Simple data class that will be used as a LINQ data source.
    public class Product
    {
        // Marked as required to satisfy the non‑nullable warning.
        public required string Name { get; set; }
        public required decimal Price { get; set; }
    }

    public static class Program
    {
        public static void Main()
        {
            // Path to the WORDML (Word 2003 XML) template file.
            // The template should contain reporting tags, e.g. <<foreach [products]>><<[Name]>> - <<[Price]:currency>><</foreach>>
            string templatePath = @"C:\Templates\ReportTemplate.xml";

            // Load the WORDML document using the Document(string) constructor.
            Document doc = new Document(templatePath);

            // Prepare a LINQ data source – a list of Product objects.
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 0.99m },
                new Product { Name = "Banana", Price = 0.59m },
                new Product { Name = "Cherry", Price = 2.49m }
            };

            // The ReportingEngine can work directly with any enumerable data source.
            // Here we expose the list under the name "products" for the template.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by merging the template with the data source.
            // The overload BuildReport(Document, object, string) allows us to name the data source.
            engine.BuildReport(doc, products, "products");

            // Save the populated document – this also verifies that the merge succeeded.
            string populatedDocPath = @"C:\Output\PopulatedReport.docx";
            doc.Save(populatedDocPath);

            // ---------------------------------------------------------------------
            // Printing
            // ---------------------------------------------------------------------
            // Aspose.Words' Document.Print() method is only available in the full .NET Framework.
            // In .NET (Core/5/6/7) the method is not present, therefore we print by saving the
            // document to a printable format (PDF) and invoking the OS print verb.
            // ---------------------------------------------------------------------
            string pdfPath = @"C:\Output\PopulatedReport.pdf";
            doc.Save(pdfPath);

            // Use the default printer via the "print" verb. This works on Windows.
            var psi = new ProcessStartInfo(pdfPath)
            {
                Verb = "print",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(psi);
        }
    }
}
