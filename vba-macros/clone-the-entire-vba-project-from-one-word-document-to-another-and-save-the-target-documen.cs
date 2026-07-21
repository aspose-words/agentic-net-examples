using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace CloneVbaProjectExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the source and target macro-enabled documents.
            const string sourcePath = "Source.docm";
            const string targetPath = "Target.docm";

            // -------------------------------------------------
            // Step 1: Create a source document with a VBA project.
            // -------------------------------------------------
            Document sourceDoc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject sourceProject = new VbaProject
            {
                Name = "SourceProject"
            };
            sourceDoc.VbaProject = sourceProject;

            // Add a procedural module with sample macro code.
            VbaModule module1 = new VbaModule
            {
                Name = "Module1",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from source\"\nEnd Sub"
            };
            sourceDoc.VbaProject.Modules.Add(module1);

            // Add a second module to demonstrate multiple modules.
            VbaModule module2 = new VbaModule
            {
                Name = "Module2",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Function Add(a As Integer, b As Integer) As Integer\n    Add = a + b\nEnd Function"
            };
            sourceDoc.VbaProject.Modules.Add(module2);

            // Save the source document in macro-enabled format.
            sourceDoc.Save(sourcePath);

            // -------------------------------------------------
            // Step 2: Load the source document and clone its VBA project.
            // -------------------------------------------------
            Document src = new Document(sourcePath);
            VbaProject clonedProject = src.VbaProject.Clone();

            // -------------------------------------------------
            // Step 3: Create a new target document and assign the cloned VBA project.
            // -------------------------------------------------
            Document targetDoc = new Document();
            targetDoc.VbaProject = clonedProject;

            // Save the target document, which now contains the cloned VBA project.
            targetDoc.Save(targetPath);
        }
    }
}
