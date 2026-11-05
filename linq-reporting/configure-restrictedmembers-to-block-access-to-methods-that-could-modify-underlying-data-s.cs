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
        // Register code page provider for Aspose.Words encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data model.
        ReportModel model = new()
        {
            Items = new List<string> { "Apple", "Banana", "Cherry" }
        };

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Configure ReportingEngine.
        ReportingEngine engine = new();

        // Block methods that could modify the data source (e.g., List<T>.Add and List<T>.Remove).
        MethodInfo? addMethod = typeof(List<string>).GetMethod("Add", new[] { typeof(string) });
        MethodInfo? removeMethod = typeof(List<string>).GetMethod("Remove", new[] { typeof(string) });

        // The RestrictedMembers property may not exist in older versions of Aspose.Words.
        // Use reflection to add the methods if the property is present.
        PropertyInfo? restrictedProp = typeof(ReportingEngine).GetProperty(
            "RestrictedMembers",
            BindingFlags.Instance | BindingFlags.Public);

        if (restrictedProp != null && restrictedProp.GetValue(engine) is System.Collections.ICollection collection)
        {
            MethodInfo? addToCollection = collection.GetType().GetMethod("Add");
            if (addToCollection != null)
            {
                if (addMethod != null) addToCollection.Invoke(collection, new object[] { addMethod });
                if (removeMethod != null) addToCollection.Invoke(collection, new object[] { removeMethod });
            }
        }

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        doc.Save(outputPath);

        // Indicate result.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Data model used by the report.
public class ReportModel
{
    public List<string> Items { get; set; } = new();
}
