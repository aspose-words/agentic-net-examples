using System;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<[model.Name]>>!");

        // Prepare the data model.
        ReportModel model = new()
        {
            Name = "World"
        };

        // Configure the reporting engine.
        ReportingEngine engine = new();

        // Attempt to configure restricted members that could write to files.
        // In newer versions of Aspose.Words the RestrictedMembers collection may be internal or removed.
        // Use reflection to add the entries if the property exists; otherwise skip silently.
        PropertyInfo? restrictedProp = typeof(ReportingEngine).GetProperty(
            "RestrictedMembers",
            BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);

        if (restrictedProp?.GetValue(null) is ICollection<string> restrictedCollection)
        {
            restrictedCollection.Add("System.IO.File.WriteAllText");
            restrictedCollection.Add("System.IO.File.WriteAllBytes");
            restrictedCollection.Add("System.IO.File.WriteAllLines");
            restrictedCollection.Add("System.IO.StreamWriter.Write");
            restrictedCollection.Add("System.IO.StreamWriter.WriteLine");
        }

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("ReportOutput.docx");
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
