using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Paths to the source document (contains VBA modules) and the target document.
        string sourcePath = @"C:\Docs\SourceWithMacros.docm";
        string targetPath = @"C:\Docs\TargetDocument.docx";
        string resultPath = @"C:\Docs\CombinedDocument.docm";

        // Load the source document that already has a VBA project.
        Document sourceDoc = new Document(sourcePath);

        // Load the target document (could be blank or existing) that will receive the VBA project.
        Document targetDoc = new Document(targetPath);

        // Create a new VBA project for the target document.
        VbaProject newProject = new VbaProject
        {
            Name = "CombinedProject"
        };
        targetDoc.VbaProject = newProject;

        // Define which module names should be copied from the source document.
        // Adjust this list as needed to select specific modules.
        HashSet<string> modulesToCopy = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Module1",
            "Module2"
        };

        // Iterate over the modules in the source document's VBA project.
        foreach (VbaModule sourceModule in sourceDoc.VbaProject.Modules)
        {
            // Copy only the modules whose names are in the selection list.
            if (modulesToCopy.Contains(sourceModule.Name))
            {
                // Clone the module to preserve its content and properties.
                VbaModule clonedModule = sourceModule.Clone();

                // Add the cloned module to the new VBA project in the target document.
                targetDoc.VbaProject.Modules.Add(clonedModule);
            }
        }

        // Save the target document with the newly added VBA project and modules.
        targetDoc.Save(resultPath);
    }
}
