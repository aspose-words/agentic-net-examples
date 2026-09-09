using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model classes
    public class Project
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
    }

    public class Department
    {
        public string Name { get; set; } = "";
        public List<Project> Projects { get; set; } = new();
    }

    public class ReportModel
    {
        public List<Department> Departments { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Outer data band – iterate over departments
            builder.Writeln("<<foreach [dept in Departments]>>");
            builder.Writeln("Department: <<[dept.Name]>>");
            builder.Writeln();

            // Inner data band – iterate over projects of the current department
            builder.Writeln("Projects:");
            builder.Writeln("<<foreach [proj in dept.Projects]>>");
            builder.Writeln("- <<[proj.Name]>>: <<[proj.Description]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // End of the outer foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare the data source
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample hierarchical data: departments with their projects
            ReportModel model = new ReportModel();
            model.Departments.Add(new Department
            {
                Name = "Human Resources",
                Projects = new List<Project>
                {
                    new Project { Name = "Recruitment", Description = "Hiring new staff members" },
                    new Project { Name = "Training", Description = "Employee development programs" }
                }
            });
            model.Departments.Add(new Department
            {
                Name = "Information Technology",
                Projects = new List<Project>
                {
                    new Project { Name = "Infrastructure", Description = "Server and network maintenance" },
                    new Project { Name = "Software Development", Description = "Internal application development" }
                }
            });

            // -------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated report
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
