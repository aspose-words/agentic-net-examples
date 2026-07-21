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

            // Create a new procedural module that contains VBA code using the Dictionary object.
            VbaModule vbaModule = new VbaModule
            {
                Name = "DictionaryModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub TestDictionary()
    ' Requires reference to Microsoft Scripting Runtime (scrrun.dll)
    Dim dict As New Scripting.Dictionary
    dict.Add ""Key1"", ""Value1""
    MsgBox dict(""Key1"")
End Sub"
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // Save the document as a macro‑enabled file.
            string outputPath = "SampleDocument.docm";
            doc.Save(outputPath, SaveFormat.Docm);

            // Indicate completion.
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
    }
}
