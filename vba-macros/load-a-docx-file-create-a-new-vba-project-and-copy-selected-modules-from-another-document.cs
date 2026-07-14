using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // File names used in the current directory.
        string sourceDocxPath = "SourceDocument.docx";
        string macroSourceDocmPath = "MacroSource.docm";
        string resultDocmPath = "ResultDocument.docm";

        // -----------------------------------------------------------------
        // Step 1: Create a blank DOCX document that will become the target.
        // -----------------------------------------------------------------
        Document targetDoc = new Document();
        targetDoc.Save(sourceDocxPath); // Saved as DOCX (no macros yet).

        // -----------------------------------------------------------------
        // Step 2: Create a macro‑enabled source document containing sample modules.
        // -----------------------------------------------------------------
        Document macroSourceDoc = new Document();

        // Create a new VBA project for the source document.
        VbaProject sourceProject = new VbaProject { Name = "SourceProject" };
        macroSourceDoc.VbaProject = sourceProject;

        // First VBA module.
        VbaModule module1 = new VbaModule
        {
            Name = "ModuleOne",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from ModuleOne!""
End Sub"
        };
        macroSourceDoc.VbaProject.Modules.Add(module1);

        // Second VBA module.
        VbaModule module2 = new VbaModule
        {
            Name = "ModuleTwo",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub ShowDate()
    MsgBox ""Today's date is "" & Date
End Sub"
        };
        macroSourceDoc.VbaProject.Modules.Add(module2);

        // Save the macro source document as a macro‑enabled .docm file.
        macroSourceDoc.Save(macroSourceDocmPath);

        // -----------------------------------------------------------------
        // Step 3: Load the target DOCX document and ensure it has a VBA project.
        // -----------------------------------------------------------------
        Document target = new Document(sourceDocxPath);

        if (target.VbaProject == null)
        {
            VbaProject newProject = new VbaProject { Name = "TargetProject" };
            target.VbaProject = newProject;
        }

        // -----------------------------------------------------------------
        // Step 4: Load the macro source document and copy selected modules.
        // -----------------------------------------------------------------
        Document source = new Document(macroSourceDocmPath);

        // Copy "ModuleOne" if it does not already exist in the target.
        VbaModule sourceModule = source.VbaProject.Modules["ModuleOne"];
        if (sourceModule != null && target.VbaProject.Modules["ModuleOne"] == null)
        {
            VbaModule clonedModule = sourceModule.Clone();
            target.VbaProject.Modules.Add(clonedModule);
        }

        // Optionally copy the second module (index 1) if present and not already in the target.
        if (source.VbaProject.Modules.Count > 1)
        {
            VbaModule secondModule = source.VbaProject.Modules[1];
            if (secondModule != null && target.VbaProject.Modules[secondModule.Name] == null)
            {
                VbaModule clonedSecond = secondModule.Clone();
                target.VbaProject.Modules.Add(clonedSecond);
            }
        }

        // -----------------------------------------------------------------
        // Step 5: Save the modified document as a macro‑enabled .docm file.
        // -----------------------------------------------------------------
        target.Save(resultDocmPath);

        // -----------------------------------------------------------------
        // Validation (optional): display macro information.
        // -----------------------------------------------------------------
        Document validationDoc = new Document(resultDocmPath);
        Console.WriteLine($"Document has macros: {validationDoc.HasMacros}");
        Console.WriteLine($"Number of VBA modules: {validationDoc.VbaProject?.Modules.Count ?? 0}");
    }
}
