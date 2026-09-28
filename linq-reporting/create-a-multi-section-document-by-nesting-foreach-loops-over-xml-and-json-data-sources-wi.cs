using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Department
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
}

public class Employee
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int DeptId { get; set; }
}

public class ReportModel
{
    public List<Department> Departments { get; set; } = new();
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // ---------- Create sample XML ----------
        string xmlPath = "departments.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Departments>
  <Department>
    <Id>1</Id>
    <Name>Human Resources</Name>
  </Department>
  <Department>
    <Id>2</Id>
    <Name>Engineering</Name>
  </Department>
</Departments>";
        File.WriteAllText(xmlPath, xmlContent);

        // ---------- Create sample JSON ----------
        string jsonPath = "employees.json";
        string jsonContent = @"[
  { ""Id"": 1, ""Name"": ""Alice"", ""DeptId"": 1 },
  { ""Id"": 2, ""Name"": ""Bob"",   ""DeptId"": 2 },
  { ""Id"": 3, ""Name"": ""Carol"", ""DeptId"": 2 },
  { ""Id"": 4, ""Name"": ""David"", ""DeptId"": 1 }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // ---------- Load data into model ----------
        var departments = XDocument.Load(xmlPath)
            .Descendants("Department")
            .Select(d => new Department
            {
                Id = (int)d.Element("Id")!,
                Name = (string)d.Element("Name")!
            })
            .ToList();

        var employees = JsonConvert.DeserializeObject<List<Employee>>(jsonContent) ?? new List<Employee>();

        var model = new ReportModel
        {
            Departments = departments,
            Employees = employees
        };

        // ---------- Build template document ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Company Report");
        builder.Writeln();

        // Outer foreach over departments
        builder.Writeln("<<foreach [dept in model.Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln();

        // Inner foreach over employees with filter
        builder.Writeln("Employees:");
        builder.Writeln("<<foreach [emp in model.Employees]>>");
        builder.Writeln("<<if [emp.DeptId == dept.Id]>>- <<[emp.Name]>> <</if>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Insert a section break after each department
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Close outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // ---------- Generate report ----------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save final document
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
