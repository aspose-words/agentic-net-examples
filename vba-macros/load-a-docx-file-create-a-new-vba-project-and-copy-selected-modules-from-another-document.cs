using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string sourceDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docm");
        string targetDocxPath = Path.Combine(Directory.GetCurrentDirectory(), "Target.docx");
        string resultDocPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docm");

        // -----------------------------------------------------------------
        // Step 1: Create a source macro-enabled document with a VBA project.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and assign it to the source document.
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
Sub GoodbyeWorld()
    MsgBox ""Goodbye from Module2!""
End Sub"
        };
        sourceProject.Modules.Add(module2);

        // Save the source document as a macro-enabled file.
        sourceDoc.Save(sourceDocPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // Step 2: Create a plain DOCX document that will receive the macros.
        // -----------------------------------------------------------------
        Document plainDoc = new Document();
        plainDoc.Save(targetDocxPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document targetDoc = new Document(targetDocxPath);

        // Ensure the target document has a VBA project; create one if absent.
        if (!targetDoc.HasMacros || targetDoc.VbaProject == null)
        {
            VbaProject targetProject = new VbaProject
            {
                Name = "TargetProject"
            };
            targetDoc.VbaProject = targetProject;
        }

        // -----------------------------------------------------------------
        // Step 3: Copy selected modules from the source document into the target.
        // -----------------------------------------------------------------
        // Load the source document again to access its VBA modules.
        Document loadedSource = new Document(sourceDocPath);

        // Example: copy only "Module1".
        VbaModule sourceModule = loadedSource.VbaProject.Modules["Module1"];
        if (sourceModule != null)
        {
            VbaModule clonedModule = sourceModule.Clone();
            // Optionally, rename the cloned module to avoid name clashes.
            clonedModule.Name = "CopiedModule1";
            targetDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Example: copy "Module2" as well.
        sourceModule = loadedSource.VbaProject.Modules["Module2"];
        if (sourceModule != null)
        {
            VbaModule clonedModule = sourceModule.Clone();
            clonedModule.Name = "CopiedModule2";
            targetDoc.VbaProject.Modules.Add(clonedModule);
        }

        // -----------------------------------------------------------------
        // Step 4: Save the resulting document as a macro-enabled file.
        // -----------------------------------------------------------------
        targetDoc.Save(resultDocPath, SaveFormat.Docm);
    }
}
