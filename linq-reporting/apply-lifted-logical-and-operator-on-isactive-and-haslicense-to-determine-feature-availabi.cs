using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<User>
            {
                new User { Name = "Alice",   IsActive = true,  HasLicense = true  },
                new User { Name = "Bob",     IsActive = true,  HasLicense = false },
                new User { Name = "Charlie", IsActive = false, HasLicense = true  },
                new User { Name = "Diana",   IsActive = null,  HasLicense = true  },
                new User { Name = "Eve",     IsActive = true,  HasLicense = null  }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }

    // Creates a Word template with LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Feature availability report");
        builder.Writeln("================================");

        // Iterate over the collection of users.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("User: <<[item.Name]>>");
        // Apply lifted logical AND on nullable booleans using explicit true comparison.
        builder.Writeln("Feature: <<if [item.IsActive == true && item.HasLicense == true]>>Enabled<</if>>");
        builder.Writeln("<<if [!(item.IsActive == true && item.HasLicense == true)]>>Disabled<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Wrapper class passed as the root data source.
public class ReportModel
{
    public List<User> Items { get; set; } = new();
}

// Data model representing a user and feature flags.
public class User
{
    public string Name { get; set; } = string.Empty;
    public bool? IsActive { get; set; }
    public bool? HasLicense { get; set; }
}
