using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names for the macro-enabled document and the output text file.
        string docPath = Path.Combine(Environment.CurrentDirectory, "MacroDocument.docm");
        string txtPath = Path.Combine(Environment.CurrentDirectory, "ModuleSource.txt");

        // -----------------------------------------------------------------
        // 1. Create a new Word document and add a VBA project with a module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleVbaProject"
        };
        doc.VbaProject = vbaProject;

        // Create a procedural VBA module, give it a name and some source code.
        VbaModule vbaModule = new VbaModule
        {
            Name = "TestModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format (.docm).
        doc.Save(docPath);

        // ---------------------------------------------------------------
        // 2. Load the document (optional – we can reuse the same instance).
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Ensure the document actually contains a VBA project.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // ---------------------------------------------------------------
        // 3. Retrieve the source code of the specific module by name.
        // ---------------------------------------------------------------
        VbaModule targetModule = loadedDoc.VbaProject.Modules["TestModule"];
        string sourceCode = targetModule?.SourceCode ?? string.Empty;

        // ---------------------------------------------------------------
        // 4. Write the retrieved source code to a text file for analysis.
        // ---------------------------------------------------------------
        File.WriteAllText(txtPath, sourceCode);
    }
}
