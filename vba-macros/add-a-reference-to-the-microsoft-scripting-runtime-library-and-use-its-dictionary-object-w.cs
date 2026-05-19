using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject project = new VbaProject
            {
                Name = "AsposeVbaProject"
            };
            doc.VbaProject = project;

            // Create a new VBA module.
            VbaModule module = new VbaModule
            {
                Name = "DictionaryModule",
                Type = VbaModuleType.ProceduralModule,
                // VBA code that uses the Dictionary object.
                // Note: This macro requires a reference to "Microsoft Scripting Runtime" (scrrun.dll).
                SourceCode = @"
' Requires reference to Microsoft Scripting Runtime
Sub UseDictionary()
    Dim dict As New Dictionary
    dict.Add ""Key1"", ""Value1""
    MsgBox dict(""Key1"")
End Sub"
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            // Save the document as a macro-enabled .docm file.
            doc.Save("DictionaryMacro.docm");
        }
    }
}
