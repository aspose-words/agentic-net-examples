using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Simple data model that will be used as the data source for the report.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }
        public bool IsFullTime { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare the data source – a list of employees.
            // -----------------------------------------------------------------
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "Alice Johnson",   Department = "Finance",   Salary = 72000m, IsFullTime = true  },
                new Employee { Name = "Bob Smith",       Department = "Finance",   Salary = 54000m, IsFullTime = false },
                new Employee { Name = "Carol Williams",  Department = "HR",        Salary = 61000m, IsFullTime = true  },
                new Employee { Name = "David Brown",     Department = "HR",        Salary = 58000m, IsFullTime = true  },
                new Employee { Name = "Eve Davis",       Department = "IT",        Salary = 95000m, IsFullTime = true  },
                new Employee { Name = "Frank Miller",    Department = "IT",        Salary = 47000m, IsFullTime = false }
            };

            // -----------------------------------------------------------------
            // 2. Use LINQ to create a hierarchical view that the template expects.
            //    The template uses an in‑table list (horizontal) with alternate
            //    content, therefore we group employees by department.
            // -----------------------------------------------------------------
            var dataSource = employees
                .GroupBy(e => e.Department)
                .Select(g => new
                {
                    Department = g.Key,
                    // The template will iterate over this collection.
                    Employees = g.Select(e => new
                    {
                        e.Name,
                        Salary = e.Salary.ToString("C"),
                        // Alternate content: Full‑time flag will be shown only for full‑time staff.
                        IsFullTime = e.IsFullTime
                    }).ToList()
                })
                .ToList();

            // -----------------------------------------------------------------
            // 3. Load the DOCX template that contains the in‑table list tags.
            // -----------------------------------------------------------------
            Document template = new Document("InTableListTemplate.docx");

            // -----------------------------------------------------------------
            // 4. Build the report using Aspose.Words ReportingEngine.
            //    The data source name ("ds") must match the name used in the template.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "ds");

            // -----------------------------------------------------------------
            // 5. Save the populated document.
            // -----------------------------------------------------------------
            template.Save("LinqReport_Output.docx");
        }
    }
}
