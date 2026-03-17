using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that contains the VBA project.
        Document srcDoc = new Document("Source.docm");

        // Create a new blank document.
        Document dstDoc = new Document();

        // Clone the entire VBA project from the source document.
        VbaProject clonedProject = srcDoc.VbaProject.Clone();

        // Assign the cloned project to the destination document.
        dstDoc.VbaProject = clonedProject;

        // Replace any modules that may already exist in the destination with the cloned ones.
        foreach (VbaModule srcModule in srcDoc.VbaProject.Modules)
        {
            // Clone the individual VBA module.
            VbaModule clonedModule = srcModule.Clone();

            // If a module with the same name already exists, remove it.
            VbaModule existing = dstDoc.VbaProject.Modules[clonedModule.Name];
            if (existing != null)
                dstDoc.VbaProject.Modules.Remove(existing);

            // Add the cloned module to the destination VBA project.
            dstDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the destination document (it will retain the macros).
        dstDoc.Save("Destination.docm");
    }
}
