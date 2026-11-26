using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template document with a LINQ Reporting tag
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Customer Name: <<[model.Name]>>");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Sample data model with a getter and setter
        MyModel model = new MyModel { Name = "John Doe" };

        // Configure ReportingEngine to allow only getter methods
        ReportingEngine engine = new ReportingEngine();

        // Use reflection to access the RestrictedMembers collection (if present) and configure it
        PropertyInfo restrictedProp = typeof(ReportingEngine).GetProperty(
            "RestrictedMembers",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        if (restrictedProp != null)
        {
            if (restrictedProp.GetValue(engine) is ICollection<MemberInfo> restrictedMembers)
            {
                restrictedMembers.Clear();

                MethodInfo nameGetter = typeof(MyModel)
                    .GetProperty(nameof(MyModel.Name))!
                    .GetGetMethod();

                if (nameGetter != null)
                {
                    restrictedMembers.Add(nameGetter);
                }
            }
        }

        // Build the report
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}

// Data model with a public property
public class MyModel
{
    public string Name { get; set; } = string.Empty;
}
