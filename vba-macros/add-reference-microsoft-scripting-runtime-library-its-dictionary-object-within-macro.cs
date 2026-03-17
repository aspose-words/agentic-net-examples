using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "MyVbaProject";
        doc.VbaProject = project;

        // Create a new procedural module that will contain the macro.
        VbaModule module = new VbaModule();
        module.Name = "DictionaryModule";
        module.Type = VbaModuleType.ProceduralModule;

        // VBA macro source code that uses the Microsoft Scripting Runtime Dictionary object.
        // The reference to the library is assumed to be added automatically when the macro runs.
        module.SourceCode = @"
Sub UseDictionary()
    Dim dict As New Dictionary
    dict.Add ""Key1"", ""Value1""
    dict.Add ""Key2"", ""Value2""

    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key & "": "" & dict(key)
    Next key
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        doc.Save("DictionaryMacro.docm");
    }
}
