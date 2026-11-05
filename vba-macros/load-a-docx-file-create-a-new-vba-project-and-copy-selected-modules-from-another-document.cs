using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a source macro-enabled document with two VBA modules.
        string sourcePath = Path.Combine(Path.GetTempPath(), "SourceMacro.docm");
        Document sourceDoc = new Document();
        sourceDoc.VbaProject = new VbaProject();

        // Add ModuleA
        string moduleACode = @"Sub Hello()
    MsgBox ""Hello from ModuleA!""
End Sub";
        VbaModule moduleA = new VbaModule();
        moduleA.Name = "ModuleA";
        moduleA.SourceCode = moduleACode;
        sourceDoc.VbaProject.Modules.Add(moduleA);

        // Add ModuleB
        string moduleBCode = @"Sub Goodbye()
    MsgBox ""Goodbye from ModuleB!""
End Sub";
        VbaModule moduleB = new VbaModule();
        moduleB.Name = "ModuleB";
        moduleB.SourceCode = moduleBCode;
        sourceDoc.VbaProject.Modules.Add(moduleB);

        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // Create a blank macro-enabled document and save it.
        string targetBlankPath = Path.Combine(Path.GetTempPath(), "BlankDocument.docm");
        Document blankDoc = new Document();
        blankDoc.VbaProject = new VbaProject(); // Ensure a VBA project exists.
        blankDoc.Save(targetBlankPath, SaveFormat.Docm);

        // Load the blank macro-enabled document.
        Document targetDoc = new Document(targetBlankPath);

        // Ensure the target document has a VBA project; create one if missing.
        if (targetDoc.VbaProject == null)
        {
            targetDoc.VbaProject = new VbaProject();
        }

        // Copy selected modules (e.g., ModuleA) from source to target.
        foreach (VbaModule srcModule in sourceDoc.VbaProject.Modules)
        {
            if (srcModule.Name.Equals("ModuleA", StringComparison.OrdinalIgnoreCase))
            {
                string sourceCode = srcModule.SourceCode ?? string.Empty;
                VbaModule copyModule = new VbaModule();
                copyModule.Name = srcModule.Name;
                copyModule.SourceCode = sourceCode;
                targetDoc.VbaProject.Modules.Add(copyModule);
            }
        }

        // Save the modified document as a macro-enabled file.
        string targetPath = Path.Combine(Path.GetTempPath(), "TargetMacro.docm");
        targetDoc.Save(targetPath, SaveFormat.Docm);

        // Validation: reload and confirm the module exists.
        Document validationDoc = new Document(targetPath);
        bool moduleFound = false;
        if (validationDoc.VbaProject != null)
        {
            foreach (VbaModule mod in validationDoc.VbaProject.Modules)
            {
                if (mod.Name.Equals("ModuleA", StringComparison.OrdinalIgnoreCase))
                {
                    moduleFound = true;
                    break;
                }
            }
        }

        Console.WriteLine($"ModuleA copied successfully: {moduleFound}");

        // Clean up temporary files.
        try { File.Delete(sourcePath); } catch { }
        try { File.Delete(targetBlankPath); } catch { }
        try { File.Delete(targetPath); } catch { }
    }
}
