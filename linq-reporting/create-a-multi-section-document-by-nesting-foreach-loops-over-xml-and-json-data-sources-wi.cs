using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings used by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        string xmlPath = "people.xml";
        File.WriteAllText(xmlPath,
@"<People>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
</People>");

        // Prepare sample JSON data.
        string jsonPath = "products.json";
        File.WriteAllText(jsonPath,
@"[
    { ""Name"": ""Laptop"", ""Price"": 1200.5 },
    { ""Name"": ""Mouse"",  ""Price"": 25.99 }
]");

        // Load XML into objects.
        List<Person> persons = new();
        XDocument xDoc = XDocument.Load(xmlPath);
        foreach (var elem in xDoc.Root?.Elements("Person") ?? [])
        {
            persons.Add(new Person
            {
                Name = (string?)elem.Element("Name") ?? string.Empty,
                Age = (int?)elem.Element("Age") ?? 0
            });
        }

        // Load JSON into objects.
        List<Product> products = JsonConvert.DeserializeObject<List<Product>>(File.ReadAllText(jsonPath)) ?? new();

        // Wrap data into a single model object.
        ReportModel model = new()
        {
            Persons = persons,
            Products = products
        };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer loop over XML persons.
        builder.Writeln("<<foreach [person in model.Persons]>>");
        builder.Writeln("Person: <<[person.Name]>> (Age: <<[person.Age]>>)");
        // Inner loop over JSON products.
        builder.Writeln("<<foreach [product in model.Products]>>");
        builder.Writeln(" - Product: <<[product.Name]>>  Price: $<<[product.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Add a section break to demonstrate multiple sections.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("All Products (outside of person loop):");
        builder.Writeln("<<foreach [product in model.Products]>>");
        builder.Writeln("Product: <<[product.Name]>> - $<<[product.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report using the model as the root data source.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
    public List<Product> Products { get; set; } = new();
}
