using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project; create one if it does not exist.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // VBA code that adds a reference to Microsoft Scripting Runtime and uses the Dictionary object.
        string vbaCode = @"
Sub AddScriptingReference()
    Dim refs As Object
    Set refs = Application.VBE.ActiveVBProject.References
    On Error Resume Next
    refs.AddFromGuid ""{420B2830-E718-11CF-893D-00A0C9054228}"", 1, 0
    On Error GoTo 0
End Sub

Sub TestDictionary()
    Dim dict As New Scripting.Dictionary
    dict.Add ""Key1"", ""Value1""
    MsgBox dict(""Key1"")
End Sub
";

        // Create a new VBA module, set its name and source code.
        VbaModule vbaModule = new VbaModule();   // parameterless constructor is supported
        vbaModule.Name = "Module1";
        vbaModule.SourceCode = vbaCode;

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format.
        const string fileName = "MacroDocument.docm";
        doc.Save(fileName, SaveFormat.Docm);

        // Reload the document and verify that the module contains the expected code.
        Document loadedDoc = new Document(fileName);
        VbaModule loadedModule = loadedDoc.VbaProject?.Modules["Module1"];
        bool containsDictionary = loadedModule?.SourceCode?.Contains("Scripting.Dictionary") ?? false;
        Console.WriteLine("Dictionary reference added: " + containsDictionary);
    }
}
