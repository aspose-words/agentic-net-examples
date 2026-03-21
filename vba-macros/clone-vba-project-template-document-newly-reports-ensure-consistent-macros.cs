using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        const string templatePath = "TemplateWithMacros.docm";
        const string outputPath = "GeneratedReport.docm";

        Document reportDoc = new Document();

        if (File.Exists(templatePath))
        {
            // Load the template document that contains the VBA project (macros).
            Document templateDoc = new Document(templatePath);

            // Clone the entire VBA project from the template.
            VbaProject clonedProject = templateDoc.VbaProject?.Clone();

            if (clonedProject != null)
            {
                // Assign the cloned VBA project to the new report document.
                reportDoc.VbaProject = clonedProject;

                // Optional: replace an existing module with the cloned one to avoid duplicates.
                VbaModule existingModule = reportDoc.VbaProject.Modules["Module1"];
                if (existingModule != null)
                {
                    VbaModule clonedModule = templateDoc.VbaProject.Modules["Module1"]?.Clone();
                    if (clonedModule != null)
                    {
                        reportDoc.VbaProject.Modules.Remove(existingModule);
                        reportDoc.VbaProject.Modules.Add(clonedModule);
                    }
                }
            }
        }
        else
        {
            Console.WriteLine($"Template file '{templatePath}' not found. Creating a report without macros.");
        }

        // Save the generated report as a macro‑enabled document.
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report saved to '{outputPath}'.");
    }
}
