using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModules
{
    static void Main()
    {
        // Load the source document (must be a macro-enabled format, e.g., .docm).
        Document srcDoc = new Document("Source.docm");

        // Ensure the document actually contains a VBA project.
        if (!srcDoc.HasMacros)
        {
            Console.WriteLine("The source document does not contain any VBA macros.");
            return;
        }

        // Get the VBA project from the source document.
        VbaProject srcProject = srcDoc.VbaProject;

        // Create a new blank document that will receive the cloned modules.
        Document destDoc = new Document();

        // Create a new VBA project for the destination document.
        VbaProject destProject = new VbaProject
        {
            Name = srcProject.Name,
            CodePage = srcProject.CodePage
        };
        destDoc.VbaProject = destProject;

        // Clone each module from the source project and add it to the destination project.
        foreach (VbaModule module in srcProject.Modules)
        {
            // Perform a deep copy of the module.
            VbaModule clonedModule = module.Clone();

            // Optionally, modify the name to avoid duplicates if needed.
            // clonedModule.Name = $"{clonedModule.Name}_Copy";

            // Add the cloned module to the destination project's collection.
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the destination document as a macro-enabled file.
        destDoc.Save("ClonedModules.docm");
        Console.WriteLine("VBA modules cloned successfully.");
    }
}
