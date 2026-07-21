using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a source document that contains a VBA project with modules.
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
            SourceCode = @"Sub Hello()
    MsgBox ""Hello from Module1""
End Sub"
        };
        sourceProject.Modules.Add(module1);

        // Create second VBA module.
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub Goodbye()
    MsgBox ""Goodbye from Module2""
End Sub"
        };
        sourceProject.Modules.Add(module2);

        // Save the source document as a macro‑enabled file.
        string sourcePath = Path.Combine(outputDir, "Source.docm");
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Create a target DOCX document (blank) and then load it.
        // -----------------------------------------------------------------
        string targetDocxPath = Path.Combine(outputDir, "Target.docx");
        Document blankDoc = new Document();
        blankDoc.Save(targetDocxPath, SaveFormat.Docx);

        // Load the DOCX document that will receive the VBA modules.
        Document targetDoc = new Document(targetDocxPath);

        // -----------------------------------------------------------------
        // 3. Ensure the target document has a VBA project.
        // -----------------------------------------------------------------
        if (targetDoc.VbaProject == null)
        {
            VbaProject targetProject = new VbaProject
            {
                Name = "TargetProject"
            };
            targetDoc.VbaProject = targetProject;
        }

        // -----------------------------------------------------------------
        // 4. Load the source document again to access its modules.
        // -----------------------------------------------------------------
        Document srcForCopy = new Document(sourcePath);
        VbaModuleCollection srcModules = srcForCopy.VbaProject.Modules;

        // -----------------------------------------------------------------
        // 5. Copy selected modules (e.g., "Module1") from source to target.
        // -----------------------------------------------------------------
        VbaModule moduleToCopy = srcModules["Module1"];
        if (moduleToCopy != null)
        {
            // Clone the module to create an independent copy.
            VbaModule copiedModule = moduleToCopy.Clone();

            // Add the cloned module to the target document's VBA project.
            targetDoc.VbaProject.Modules.Add(copiedModule);
        }

        // -----------------------------------------------------------------
        // 6. Save the target document as a macro‑enabled file.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.docm");
        targetDoc.Save(resultPath, SaveFormat.Docm);

        // Simple validation (optional): ensure the macro was copied.
        Document validationDoc = new Document(resultPath);
        Console.WriteLine($"Has macros: {validationDoc.HasMacros}");
        Console.WriteLine($"Modules count: {validationDoc.VbaProject?.Modules?.Count ?? 0}");
        VbaModule copied = validationDoc.VbaProject?.Modules["Module1"];
        Console.WriteLine($"Copied module source code:\n{copied?.SourceCode}");
    }
}
