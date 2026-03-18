using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the template document that contains the VBA project (macros).
        Document templateDoc = new Document("TemplateWithMacros.docm");

        // Create a new blank document that will become the report.
        Document reportDoc = new Document();

        // Clone the entire VBA project from the template.
        VbaProject clonedProject = templateDoc.VbaProject.Clone();

        // Assign the cloned VBA project to the new report document.
        reportDoc.VbaProject = clonedProject;

        // If the report already contains a module with the same name as one in the template,
        // replace it with the cloned module to avoid duplicates.
        // (This step is optional for a completely blank document.)
        VbaModule existingModule = reportDoc.VbaProject.Modules["Module1"];
        if (existingModule != null)
        {
            // Clone the specific module from the template.
            VbaModule clonedModule = templateDoc.VbaProject.Modules["Module1"].Clone();

            // Remove the existing module and add the cloned one.
            reportDoc.VbaProject.Modules.Remove(existingModule);
            reportDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the generated report as a macro‑enabled document.
        reportDoc.Save("GeneratedReport.docm");
    }
}
