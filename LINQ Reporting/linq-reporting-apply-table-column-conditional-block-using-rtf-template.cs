using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Prepare the data source (LINQ-friendly object)
        // -------------------------------------------------
        var data = new
        {
            Products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20, InStock = true  },
                new Product { Name = "Banana", Price = 0.80, InStock = false },
                new Product { Name = "Orange", Price = 1.50, InStock = true  }
            }
        };

        // -------------------------------------------------
        // 2. Define an RTF template that contains a table.
        //    The table iterates over the Products collection.
        //    For each product we display the Name column always.
        //    The Price column is shown only when InStock == true
        //    using a conditional block.
        // -------------------------------------------------
        string rtfTemplate = @"{\rtf1\ansi
{\trowd\cellx2000\cellx4000
\intbl <<foreach [Products]>>\intbl <<[Name]>>\cell 
\intbl <<if [InStock]>><<[Price]>>\cell<<endif>>
\intbl <<endforeach>>\row}
}";

        // -------------------------------------------------
        // 3. Load the RTF template into an Aspose.Words Document.
        //    (Using the provided load rule – a MemoryStream here.)
        // -------------------------------------------------
        Document doc;
        using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(rtfTemplate)))
        {
            doc = new Document(stream);
        }

        // -------------------------------------------------
        // 4. Build the report using ReportingEngine.
        //    RemoveEmptyParagraphs ensures that rows where the
        //    conditional block is not rendered do not leave empty cells.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, data, "data");

        // -------------------------------------------------
        // 5. Save the populated document (using the provided save rule).
        // -------------------------------------------------
        doc.Save("Report.rtf");
    }

    // Simple POCO representing a product.
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public bool InStock { get; set; }
    }
}
