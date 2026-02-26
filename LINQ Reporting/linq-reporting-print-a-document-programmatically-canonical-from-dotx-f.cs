using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOTX template from file.
        Document doc = new Document("Template.dotx");

        // Sample LINQ data source – a list of products.
        var products = new List<Product>
        {
            new Product("Apple", 1.20),
            new Product("Banana", 0.80),
            new Product("Cherry", 2.50)
        };

        // Populate the template using the ReportingEngine.
        // The data source name ("Products") can be referenced in the template as <<foreach [Products]>>
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, products, "Products");

        // ---------------------------------------------------------------------
        // Printing – Aspose.Words provides Document.Print only on .NET Framework.
        // For .NET Core / .NET 5+ we save the document to a printable format (PDF)
        // and invoke the OS print command.
        // ---------------------------------------------------------------------
        string outPath = Path.Combine(Path.GetTempPath(), "Report.pdf");
        doc.Save(outPath, SaveFormat.Pdf);

        // Use the default printer via the OS "print" verb.
        var psi = new ProcessStartInfo(outPath)
        {
            Verb = "print",
            CreateNoWindow = true,
            UseShellExecute = true
        };
        Process.Start(psi);
    }

    // Simple POCO class used as the LINQ data source.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }

        public Product() { }
        public Product(string name, double price)
        {
            Name = name;
            Price = price;
        }
    }
}
