using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Email { get; set; } = "";
    
    public Person(string name, int age, string email)
    {
        Name = name;
        Age = age;
        Email = email;
    }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        // Sample data.
        Person person = new("John Doe", 30, "john.doe@example.com");
        
        // Create template document.
        Document template = new();
        DocumentBuilder builder = new(template);
        
        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        
        // Save template to disk (required before building the report).
        string templatePath = "Template.docx";
        template.Save(templatePath);
        
        // Load the template back.
        Document doc = new(templatePath);
        
        // Build the report.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, person, "person");
        
        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
        
        // Indicate completion.
        Console.WriteLine(success
            ? $"Report generated successfully: {Path.GetFullPath(outputPath)}"
            : "Report generation failed.");
    }
}
