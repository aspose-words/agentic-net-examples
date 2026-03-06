using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Vba;

namespace AsposeWordsMacroDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX file. The format is detected automatically.
            Document doc = new Document("InputDocument.docx");

            // -----------------------------------------------------------------
            // 1. Add a VBA project with a simple macro to the document.
            // -----------------------------------------------------------------
            // Create a new VBA project and assign it to the document.
            VbaProject vbaProject = new VbaProject
            {
                Name = "AsposeDemoProject"
            };
            doc.VbaProject = vbaProject;

            // Create a new module that will contain the macro source code.
            VbaModule vbaModule = new VbaModule
            {
                Name = "Module1",
                Type = VbaModuleType.ProceduralModule,
                // Simple macro that shows a message box.
                SourceCode = @"
Sub ShowMessage()
    MsgBox ""Hello from Aspose.Words macro!""
End Sub"
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // -----------------------------------------------------------------
            // 2. Insert a MACROBUTTON field that runs the macro when double‑clicked.
            // -----------------------------------------------------------------
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Move the cursor to the end of the document.
            builder.MoveToDocumentEnd();

            // Insert a paragraph break before the field for readability.
            builder.Writeln();

            // Insert the MACROBUTTON field.
            FieldMacroButton macroButton = (FieldMacroButton)builder.InsertField(FieldType.FieldMacroButton, true);
            macroButton.MacroName = "ShowMessage";
            macroButton.DisplayText = "Double‑click to run macro";

            // -----------------------------------------------------------------
            // 3. Save the document as a macro‑enabled file (.docm).
            // -----------------------------------------------------------------
            doc.Save("DocumentWithMacro.docm");

            // -----------------------------------------------------------------
            // 4. Demonstrate removal of macros from a document.
            // -----------------------------------------------------------------
            // Load the macro‑enabled document we just created.
            Document docWithMacro = new Document("DocumentWithMacro.docm");

            // Verify that the document contains macros.
            Console.WriteLine($"Has macros before removal: {docWithMacro.HasMacros}");

            // Remove all macros, toolbars, and customizations.
            docWithMacro.RemoveMacros();

            // Verify removal.
            Console.WriteLine($"Has macros after removal: {docWithMacro.HasMacros}");

            // Save the cleaned document as a regular DOCX file.
            docWithMacro.Save("DocumentWithoutMacro.docx");
        }
    }
}
