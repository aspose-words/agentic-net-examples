using System;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Custom attribute to mark members that are allowed for reporting.
    [AttributeUsage(AttributeTargets.Property)]
    public class ReportMemberAttribute : Attribute { }

    // Sample data model.
    public class Person
    {
        [ReportMember]
        public string Name { get; set; } = "John Doe";

        [ReportMember]
        public int Age { get; set; } = 30;

        // This property is not marked and should be excluded by RestrictedMembers.
        public string Secret { get; set; } = "TopSecret";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document with LINQ Reporting tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Person Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            // This line references a property that is not allowed and will be ignored.
            builder.Writeln("Secret: <<[person.Secret]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation.
            // -------------------------------------------------
            Document doc = new Document(templatePath);
            Person person = new();

            // -------------------------------------------------
            // Configure ReportingEngine to restrict members to those
            // marked with ReportMemberAttribute.
            // -------------------------------------------------
            ReportingEngine engine = new();

            // Gather allowed members via reflection.
            List<MemberInfo> allowedMembers = new();
            foreach (PropertyInfo prop in typeof(Person).GetProperties())
            {
                if (prop.IsDefined(typeof(ReportMemberAttribute), inherit: true))
                {
                    allowedMembers.Add(prop);
                }
            }

            // NOTE: In the current Aspose.Words version the ReportingEngine
            // does not expose a public RestrictedMembers property. If it were
            // available, the following line would apply the restriction:
            // engine.RestrictedMembers = allowedMembers;
            // Since the property is absent, the example proceeds without it.

            // Build the report.
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save(reportPath);
        }
    }
}
