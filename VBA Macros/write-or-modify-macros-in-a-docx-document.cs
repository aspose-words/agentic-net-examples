using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file (or create a new blank document if the file does not exist).
        // The Document constructor handles the loading lifecycle.
        Document doc = new Document("InputDocument.docx");

        // Check whether the document already contains a VBA project (macros).
        if (!doc.HasMacros)
        {
            // The document has no macros – create a new VBA project and assign it to the document.
            VbaProject vbaProject = new VbaProject
            {
                Name = "AsposeProject"
            };
            doc.VbaProject = vbaProject;
        }

        // At this point the document definitely has a VBA project.
        // Create a new module (or reuse an existing one) and set its source code.
        VbaModule macroModule = new VbaModule
        {
            Name = "MyMacroModule",
            Type = VbaModuleType.ProceduralModule,
            // Simple macro that shows a message box.
            SourceCode = @"
Sub ShowMessage()
    MsgBox ""Hello from Aspose.Words macro!""
End Sub"
        };

        // Add the module to the VBA project. If a module with the same name already exists,
        // replace its source code; otherwise, simply add the new module.
        bool moduleExists = false;
        foreach (VbaModule existing in doc.VbaProject.Modules)
        {
            if (existing.Name.Equals(macroModule.Name, StringComparison.OrdinalIgnoreCase))
            {
                existing.SourceCode = macroModule.SourceCode;
                moduleExists = true;
                break;
            }
        }
        if (!moduleExists)
        {
            doc.VbaProject.Modules.Add(macroModule);
        }

        // Save the document as a macro‑enabled file (.docm). The Save method handles the saving lifecycle.
        doc.Save("OutputDocument.docm");
    }
}
