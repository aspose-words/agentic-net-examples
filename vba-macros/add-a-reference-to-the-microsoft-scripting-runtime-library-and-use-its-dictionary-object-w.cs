using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "AsposeProject"
        };
        doc.VbaProject = project;

        // Create a new procedural VBA module.
        VbaModule module = new VbaModule
        {
            Name = "DictionaryMacro",
            Type = VbaModuleType.ProceduralModule,
            // VBA code that adds a reference to Microsoft Scripting Runtime (scrrun.dll)
            // and uses the Dictionary object.
            SourceCode = @"
Sub AddReferenceAndUseDictionary()
    Dim refs As Object
    Set refs = Application.VBE.ActiveVBProject.References
    refs.AddFromFile ""C:\Windows\System32\scrrun.dll""
    Dim dict As New Dictionary
    dict.Add ""Key1"", ""Value1""
    MsgBox dict(""Key1"")
End Sub
"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        doc.Save("DictionaryMacro.docm");
    }
}
