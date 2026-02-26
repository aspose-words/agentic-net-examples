using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModuleExample
{
    static void Main()
    {
        // Path to the source document that already contains VBA macros (e.g., a .docm file).
        const string sourcePath = @"C:\Docs\Source.docm";

        // Path to the destination document (a .docx file without macros).
        const string destinationPath = @"C:\Docs\Destination.docx";

        // Path where the resulting document (with the cloned VBA module) will be saved.
        const string outputPath = @"C:\Docs\Result.docm";

        // Load the source document.
        Document srcDoc = new Document(sourcePath);

        // Retrieve the VBA module you want to clone.
        // Here we clone the first module in the collection; adjust the index or name as needed.
        VbaModule originalModule = srcDoc.VbaProject.Modules[0];

        // Perform a deep clone of the module.
        VbaModule clonedModule = originalModule.Clone();

        // Load the destination document.
        Document destDoc = new Document(destinationPath);

        // Ensure the destination document has a VBA project; create one if it does not.
        if (destDoc.VbaProject == null)
        {
            destDoc.VbaProject = new VbaProject();
            destDoc.VbaProject.Name = "ClonedProject";
        }

        // Add the cloned module to the destination document's VBA project.
        destDoc.VbaProject.Modules.Add(clonedModule);

        // Save the result as a macro-enabled document (.docm) so the VBA code is retained.
        destDoc.Save(outputPath);
    }
}
