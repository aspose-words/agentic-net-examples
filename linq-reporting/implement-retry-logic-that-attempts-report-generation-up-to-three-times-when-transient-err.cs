using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document programmatically
        CreateTemplate(templatePath);

        // Load the template
        Document templateDoc = new Document(templatePath);

        // Prepare sample data
        Order order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" }
            }
        };

        // Initialize the reporting engine
        ReportingEngine engine = new ReportingEngine();

        bool success = false;
        const int maxAttempts = 3;

        for (int attempt = 1; attempt <= maxAttempts && !success; attempt++)
        {
            try
            {
                // Build the report
                engine.BuildReport(templateDoc, order, "order");
                success = true;
            }
            catch (Exception ex) when (IsTransient(ex))
            {
                // Transient error encountered; log and retry
                Console.WriteLine($"Attempt {attempt} failed with a transient error: {ex.Message}");
                if (attempt == maxAttempts)
                {
                    Console.WriteLine("Maximum retry attempts reached. Report generation failed.");
                }
            }
        }

        if (success)
        {
            // Save the generated report
            templateDoc.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }

    // Determines whether an exception is considered transient.
    private static bool IsTransient(Exception ex)
    {
        // For demonstration, treat all IOExceptions as transient.
        // Extend this method to include other transient scenarios as needed.
        return ex is IOException;
    }

    // Creates a simple Word template with LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a heading
        builder.Writeln("Order Report");
        builder.Writeln();

        // Insert a tag for the customer's name
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin foreach loop for items
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table header
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.EndRow();

        // Table row for each item
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // End table and foreach
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(path);
    }
}

// Data model classes
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
