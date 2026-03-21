using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

[AttributeUsage(AttributeTargets.Property)]
public class ReportAllowedAttribute : Attribute { }

public class Person
{
    // This property will be accessible in the template.
    [ReportAllowed]
    public string Name { get; set; }

    // This property will be blocked by the reporting engine.
    public int Age { get; set; }
}

public class ReportData
{
    public Person person { get; set; }
}

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple template document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Name: <<[person.Name]>>"); // allowed
        builder.Writeln("Age: <<[person.Age]>>");   // should be blocked
        builder.Writeln(); // add a blank line for readability

        // -----------------------------------------------------------------
        // 2. Determine which types need to be restricted.
        //    Any type that contains at least one property NOT marked with
        //    ReportAllowedAttribute will be added to the restricted list.
        // -----------------------------------------------------------------
        Type[] typesToInspect = { typeof(Person) };
        List<Type> restrictedTypes = new List<Type>();

        foreach (Type t in typesToInspect)
        {
            bool hasDisallowedMember = t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                        .Any(p => p.GetCustomAttribute<ReportAllowedAttribute>() == null);
            if (hasDisallowedMember)
                restrictedTypes.Add(t);
        }

        // -----------------------------------------------------------------
        // 3. Apply the restriction BEFORE the first report build.
        // -----------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(restrictedTypes.ToArray());

        // -----------------------------------------------------------------
        // 4. Build the report.
        //    AllowMissingMembers prevents an exception when a blocked member
        //    is referenced; the engine will output an empty string instead.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        var data = new ReportData
        {
            person = new Person
            {
                Name = "John Doe",
                Age = 42
            }
        };

        engine.BuildReport(doc, data);

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save("Report.docx");
    }
}
