using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a blank Word document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // 2. Insert LINQ Reporting tags into the template.
        //    <<[customer.Name]>> will be replaced with the Name property.
        //    <<[customer.Total]:dollarText>> will format the Total as dollar text.
        builder.Writeln("Customer: <<[customer.Name]>>");
        builder.Writeln("Total: <<[customer.Total]:dollarText>>");

        // 3. Prepare the data source object.
        var customer = new Customer
        {
            Name = "Acme Corp",
            Total = 12345.67M
        };

        // 4. Populate the template using ReportingEngine.
        //    The overload with a dataSourceName allows the template to reference the object itself.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, customer, "customer");

        // 5. Save the resulting document as a PostScript (PS) file.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps,
            // Example of an additional option; can be omitted if defaults are sufficient.
            UseHighQualityRendering = true
        };
        template.Save("Report.ps", psOptions);
    }

    // Simple POCO class used as the data source for the report.
    public class Customer
    {
        public string Name { get; set; }
        public decimal Total { get; set; }
    }
}
