using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -------------------- Add VBA project with a macro --------------------
        // The macro copies the current selection to the Windows clipboard.
        const string macroCode = @"
Sub CopySelection()
    Selection.Copy
End Sub
";

        // Initialize a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "AsposeProject"
        };
        doc.VbaProject = vbaProject;

        // Create a procedural module that contains the macro.
        VbaModule module = new VbaModule
        {
            Name = "CopyModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = macroCode
        };
        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // -------------------- Insert a MACROBUTTON field --------------------
        // This field will appear as a clickable button that runs the macro.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldMacroButton macroButton = (FieldMacroButton)builder.InsertField(FieldType.FieldMacroButton, true);
        macroButton.MacroName = "CopySelection";               // Name of the macro to run.
        macroButton.DisplayText = "Copy selection to clipboard"; // Text shown as the button.

        // -------------------- Save the document --------------------
        // Save as a macro‑enabled document (.docm) so the VBA code is retained.
        doc.Save("CopySelectionMacro.docm");
    }
}
