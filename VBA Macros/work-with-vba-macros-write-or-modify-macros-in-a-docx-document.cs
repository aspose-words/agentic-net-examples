using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaMacroExample
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new blank document and add a VBA project with a macro.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // Create a blank document
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains a VBA macro."); // Add some visible text

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "MyProject";
        doc.VbaProject = project;

        // Create a new procedural module that holds the macro code.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = 
@"Sub MyMacro()
    MsgBox ""Hello from macro!""
End Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file (.docm).
        doc.Save("MyMacroDoc.docm");

        // -----------------------------------------------------------------
        // 2. Load the previously saved document, modify the macro, and save again.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("MyMacroDoc.docm"); // Load the macro‑enabled document

        // Verify that the document indeed contains macros.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            // Access the module by name and change its source code.
            VbaModule existingModule = loadedDoc.VbaProject.Modules["MyModule"];
            existingModule.SourceCode = 
@"Sub MyMacro()
    MsgBox ""Macro has been modified!""
End Sub";

            // Save the modified document under a new name.
            loadedDoc.Save("MyMacroDocModified.docm");
        }
        else
        {
            Console.WriteLine("The loaded document does not contain any VBA macros.");
        }
    }
}
