using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Step 1: Create a source macro-enabled document with a VBA project
            // and a couple of modules that we will copy later.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();

            // Create a new VBA project for the source document.
            VbaProject sourceProject = new VbaProject
            {
                Name = "SourceProject"
            };
            sourceDoc.VbaProject = sourceProject;

            // Create first VBA module.
            VbaModule module1 = new VbaModule
            {
                Name = "Module1",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Module1!""
End Sub"
            };
            sourceProject.Modules.Add(module1);

            // Create second VBA module.
            VbaModule module2 = new VbaModule
            {
                Name = "Module2",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub ShowDate()
    MsgBox ""Current date: "" & Date
End Sub"
            };
            sourceProject.Modules.Add(module2);

            // Save the source document as a macro‑enabled file.
            string sourcePath = Path.Combine(outputDir, "Source.docm");
            sourceDoc.Save(sourcePath, SaveFormat.Docm);

            // -----------------------------------------------------------------
            // Step 2: Create a target DOCX document (without macros) that we will
            // load, attach a new VBA project, and copy selected modules into.
            // -----------------------------------------------------------------
            Document targetDoc = new Document(); // blank DOCX document
            string targetDocxPath = Path.Combine(outputDir, "Target.docx");
            targetDoc.Save(targetDocxPath, SaveFormat.Docx);

            // Load the DOCX document.
            Document loadedTarget = new Document(targetDocxPath);

            // Ensure the target document has a VBA project; create one if missing.
            if (loadedTarget.VbaProject == null)
            {
                VbaProject newProject = new VbaProject
                {
                    Name = "TargetProject"
                };
                loadedTarget.VbaProject = newProject;
            }

            // -----------------------------------------------------------------
            // Step 3: Load the source document again and copy selected modules.
            // -----------------------------------------------------------------
            Document loadedSource = new Document(sourcePath);
            VbaProject sourceVba = loadedSource.VbaProject;

            // Choose which modules to copy (by name).
            string[] modulesToCopy = { "Module1", "Module2" };

            foreach (string moduleName in modulesToCopy)
            {
                VbaModule srcModule = sourceVba.Modules[moduleName];
                if (srcModule != null)
                {
                    // Clone the module to avoid reference issues.
                    VbaModule clonedModule = srcModule.Clone();

                    // Add the cloned module to the target's VBA project.
                    loadedTarget.VbaProject.Modules.Add(clonedModule);
                }
            }

            // -----------------------------------------------------------------
            // Step 4: Save the modified target document as a macro‑enabled file.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(outputDir, "TargetWithMacros.docm");
            loadedTarget.Save(resultPath, SaveFormat.Docm);

            // Simple validation: output the number of modules in the target document.
            Console.WriteLine($"Target document now contains {loadedTarget.VbaProject.Modules.Count} VBA modules.");
        }
    }
}
