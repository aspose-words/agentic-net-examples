using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace LinqReportingRestrictedMembers
{
    // Custom attribute to mark members that are allowed in the report.
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ReportMemberAttribute : Attribute { }

    // Original data model containing all properties.
    public class ReportModel
    {
        [ReportMember]
        public string Name { get; set; } = "John Doe";

        public int Age { get; set; } = 30;

        [ReportMember]
        public string City { get; set; } = "New York";

        public string Secret { get; set; } = "TopSecret";
    }

    // Wrapper model exposing only the members marked with ReportMemberAttribute.
    public class RestrictedReportModel
    {
        public string Name { get; set; }
        public string City { get; set; }

        public RestrictedReportModel(ReportModel source)
        {
            // Copy only the allowed properties.
            Name = source.Name;
            City = source.City;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document with LINQ Reporting tags.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            CreateTemplate(templatePath);

            // 2. Load the template.
            Document doc = new Document(templatePath);

            // 3. Prepare the full data model.
            ReportModel fullModel = new();

            // 4. Create a restricted model that contains only the allowed members.
            RestrictedReportModel restrictedModel = new(fullModel);

            // 5. Build the report using the restricted model.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, restrictedModel, "model");

            // 6. Save the generated report.
            string resultPath = Path.Combine(outputDir, "Result.docx");
            doc.Save(resultPath);
        }

        private static void CreateTemplate(string path)
        {
            Document doc = new();
            DocumentBuilder builder = new(doc);

            // Write a simple paragraph with allowed properties.
            builder.Writeln("Customer Report");
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("City: <<[model.City]>>");

            // The following lines are intentionally omitted because the corresponding
            // properties are not exposed in the restricted model.
            // builder.Writeln("Age: <<[model.Age]>>");
            // builder.Writeln("Secret: <<[model.Secret]>>");

            doc.Save(path);
        }
    }
}
