using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class CloneVbaProjectExample
{
    public static void Main()
    {
        // Create a source macro-enabled document.
        Document sourceDoc = new Document();

        // Create a new VBA project and assign it to the source document.
        VbaProject sourceProject = new VbaProject();
        sourceProject.Name = "SourceProject";
        sourceDoc.VbaProject = sourceProject;

        // Create a VBA module with some simple code.
        VbaModule sourceModule = new VbaModule();
        sourceModule.Name = "Module1";
        sourceModule.Type = VbaModuleType.ProceduralModule;
        sourceModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from source document!""
End Sub";
        // Add the module to the VBA project.
        sourceDoc.VbaProject.Modules.Add(sourceModule);

        // Save the source document (optional, just for demonstration).
        sourceDoc.Save("Source.docm");

        // Clone the VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a destination document.
        Document destDoc = new Document();

        // Assign the cloned VBA project to the destination document.
        destDoc.VbaProject = clonedProject;

        // The destination document now contains a default module (created during cloning).
        // Replace it with the cloned module from the source to preserve original module order.
        VbaModule oldModule = destDoc.VbaProject.Modules["Module1"];
        if (oldModule != null)
        {
            // Clone the original module from the source document.
            VbaModule copiedModule = sourceDoc.VbaProject.Modules["Module1"].Clone();

            // Remove the default module and add the copied one.
            destDoc.VbaProject.Modules.Remove(oldModule);
            destDoc.VbaProject.Modules.Add(copiedModule);
        }

        // Save the destination document with the duplicated VBA project.
        destDoc.Save("ClonedVbaProject.docm");
    }
}
