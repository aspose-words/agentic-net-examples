using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Groups = new()
            {
                new OrderGroup
                {
                    GroupName = "Electronics",
                    Orders = new()
                    {
                        new Order { Name = "Smartphone", Quantity = 5 },
                        new Order { Name = "Laptop", Quantity = 2 }
                    }
                },
                new OrderGroup
                {
                    GroupName = "Books",
                    Orders = new()
                    {
                        new Order { Name = "C# in Depth", Quantity = 3 },
                        new Order { Name = "ASP.NET Core Guide", Quantity = 4 }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach over groups.
        builder.Writeln("<<foreach [group in Model.Groups]>>");
        // Group name.
        builder.Writeln("<<[group.GroupName]>>");
        builder.Writeln();

        // Start a numbered list for the orders of the current group.
        builder.ListFormat.List = template.Lists.Add(ListTemplate.NumberDefault);

        // Restart numbering for each group, then iterate over orders.
        builder.Writeln("<<restartNum>><<foreach [order in group.Orders]>>" +
                        "<<[order.Name]>> - <<[order.Quantity]>>" +
                        "<</foreach>>");

        // End the list for this group.
        builder.ListFormat.RemoveNumbers();

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // The root object name in the template is "Model".
        bool success = engine.BuildReport(doc, model, "Model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Optional: indicate success (no console interaction required).
        // In a real scenario you might log this information.
        if (!success)
        {
            throw new InvalidOperationException("Report generation failed.");
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<OrderGroup> Groups { get; set; } = new();
}

public class OrderGroup
{
    public string GroupName { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}
