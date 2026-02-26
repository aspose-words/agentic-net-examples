using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModule
{
    static void Main()
    {
        // Load the source document that contains VBA macros.
        Document doc = new Document("Input.docm");

        // Access the VBA project of the document.
        VbaProject vbaProject = doc.VbaProject;

        // Retrieve the module you want to clone (by name or index).
        VbaModule sourceModule = vbaProject.Modules["Module1"]; // adjust the name as needed

        // Perform a deep clone of the module.
        VbaModule clonedModule = sourceModule.Clone();

        // Give the cloned module a unique name.
        clonedModule.Name = "Module1_Copy";

        // Add the cloned module back into the VBA project.
        vbaProject.Modules.Add(clonedModule);

        // Save the document with the new cloned module.
        doc.Save("Output.docm");
    }
}
