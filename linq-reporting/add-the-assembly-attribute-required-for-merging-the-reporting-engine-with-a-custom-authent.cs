using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

[assembly: ReportingEngineAuthenticationAttribute(typeof(MyAuthModule))]

[AttributeUsage(AttributeTargets.Assembly)]
public sealed class ReportingEngineAuthenticationAttribute : Attribute
{
    public Type AuthModuleType { get; }

    public ReportingEngineAuthenticationAttribute(Type authModuleType) => AuthModuleType = authModuleType;
}

public class MyAuthModule
{
    // Dummy authentication logic for illustration.
    public bool Authenticate(string user) => user == "admin";
}

public class Model
{
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // Register code pages provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple template document with a LINQ Reporting tag.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello, <<[model.Name]>>!");

        // Save the template to disk (optional, demonstrates file handling).
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        var model = new Model();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
