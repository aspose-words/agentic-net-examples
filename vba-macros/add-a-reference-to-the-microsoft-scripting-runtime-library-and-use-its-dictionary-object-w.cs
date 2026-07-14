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
            VbaProject vbaProject = new VbaProject
            {
                Name = "SampleProject"
            };
            doc.VbaProject = vbaProject;

            // Create a new procedural VBA module.
            VbaModule vbaModule = new VbaModule
            {
                Name = "DictionaryModule",
                Type = VbaModuleType.ProceduralModule,
                // VBA macro that uses the Scripting.Dictionary object.
                // The Microsoft Scripting Runtime library should be referenced in the VBA project.
                SourceCode = @"
Sub UseDictionary()
    ' Create a Dictionary object (requires reference to Microsoft Scripting Runtime).
    Dim dict As Object
    Set dict = CreateObject(""Scripting.Dictionary"")
    dict.Add ""Key1"", ""Value1""
    MsgBox dict(""Key1"")
End Sub"
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // Save the document as a macro‑enabled file.
            doc.Save("DictionaryMacro.docm");
        }
    }
}
