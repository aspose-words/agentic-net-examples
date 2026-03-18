using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HierarchicalReportExample
{
    // Data model classes
    public class Project
    {
        public string Name { get; set; }
        public double Budget { get; set; }
    }

    public class Department
    {
        public string Name { get; set; }
        public List<Project> Projects { get; set; } = new List<Project>();
    }

    public class Company
    {
        public List<Department> Departments { get; set; } = new List<Department>();
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document programmatically.
            Document doc = new Document();                     // create
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Company Projects Report");
            builder.Writeln();

            // Outer data band – iterate over departments.
            builder.Writeln("<<foreach [Departments]>>");
            builder.Writeln("Department: <<[Name]>>");
            builder.Writeln("Projects:");

            // Inner data band – iterate over projects of the current department.
            builder.Writeln("<<foreach [Projects]>>");
            builder.Writeln("- <<[Name]>> (Budget: <<[Budget]:dollarText>>)");
            builder.Writeln("<</foreach>>"); // end inner foreach

            builder.Writeln("<</foreach>>"); // end outer foreach

            // 2. Build the hierarchical data source.
            var company = new Company
            {
                Departments = new List<Department>
                {
                    new Department
                    {
                        Name = "Human Resources",
                        Projects = new List<Project>
                        {
                            new Project { Name = "Recruitment", Budget = 5000 },
                            new Project { Name = "Training", Budget = 3000 }
                        }
                    },
                    new Department
                    {
                        Name = "Information Technology",
                        Projects = new List<Project>
                        {
                            new Project { Name = "Infrastructure Upgrade", Budget = 20000 },
                            new Project { Name = "Software Development", Budget = 15000 }
                        }
                    }
                }
            };

            // 3. Populate the template using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, company); // build report

            // 4. Save the final document.
            doc.Save("HierarchicalReport.docx"); // save
        }
    }
}
