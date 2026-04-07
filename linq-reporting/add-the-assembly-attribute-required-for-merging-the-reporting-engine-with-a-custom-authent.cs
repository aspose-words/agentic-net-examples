using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

// Assembly attribute that merges the reporting engine with a custom authentication module.
[assembly: Aspose.Words.Reporting.CustomAuthenticationModule(typeof(MyAuthModule))]

namespace Aspose.Words.Reporting
{
    // Dummy attribute definition – in real scenarios Aspose.Words provides this.
    [AttributeUsage(AttributeTargets.Assembly, AllowMultiple = false)]
    public sealed class CustomAuthenticationModuleAttribute : Attribute
    {
        public Type ModuleType { get; }

        public CustomAuthenticationModuleAttribute(Type moduleType)
        {
            ModuleType = moduleType;
        }
    }
}

// Simple custom authentication module placeholder.
public class MyAuthModule
{
    // In a real implementation this would contain authentication logic.
    public bool Authenticate(string resource) => true;
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    public string Greeting { get; set; } = "Hello from LINQ Reporting!";
}

// Console application entry point.
public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a Word template containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Report:");
        builder.Writeln("<<[model.Greeting]>>"); // LINQ Reporting tag.
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // 3. Build the report using Aspose.Words.ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Indicate successful completion.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
