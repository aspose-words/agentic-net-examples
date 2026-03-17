using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleCopyExample
{
    static void Main()
    {
        // Paths to the source and destination documents.
        string dataDir = @"C:\Data\";
        string sourcePath = dataDir + "Source.docm";
        string destinationPath = dataDir + "Destination.docm";
        string resultPath = dataDir + "Result.docm";

        // Load the source and destination documents.
        Document srcDoc = new Document(sourcePath);
        Document dstDoc = new Document(destinationPath);

        // Ensure the destination document has a VBA project.
        if (dstDoc.VbaProject == null)
            dstDoc.VbaProject = new VbaProject();

        // Name of the VBA module to copy.
        string moduleName = "Module1";

        // Retrieve the module from the source document.
        VbaModule srcModule = srcDoc.VbaProject.Modules[moduleName];
        if (srcModule != null)
        {
            // Clone the source module.
            VbaModule clonedModule = srcModule.Clone();

            // If a module with the same name already exists in the destination, remove it.
            VbaModule existingModule = dstDoc.VbaProject.Modules[moduleName];
            if (existingModule != null)
                dstDoc.VbaProject.Modules.Remove(existingModule);

            // Add the cloned module to the destination document's VBA project.
            dstDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the modified destination document.
        dstDoc.Save(resultPath);
    }
}
