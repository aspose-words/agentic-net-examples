using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
            doc.VbaProject = new VbaProject();

        // Add a VBA module that uses the Dictionary object from the Scripting Runtime.
        VbaModuleCollection modules = doc.VbaProject.Modules;
        const string moduleName = "Module1";
        string vbaCode = @"Option Explicit
Sub UseDictionary()
    Dim dict As New Dictionary
    dict.Add ""Key1"", ""Value1""
    MsgBox dict(""Key1"")
End Sub";

        // Add the module if it does not already exist.
        if (modules[moduleName] == null)
        {
            // Create a new module, set its name and source code, then add it to the collection.
            VbaModule newModule = new VbaModule();
            newModule.Name = moduleName;
            newModule.SourceCode = vbaCode;
            modules.Add(newModule);
        }

        // Save the document in a macro‑enabled format.
        const string fileName = "MacroDocument.docm";
        doc.Save(fileName, SaveFormat.Docm);

        // Reload the document to verify that the module was added.
        Document loadedDoc = new Document(fileName);
        VbaProject loadedVba = loadedDoc.VbaProject;

        bool moduleFound = loadedVba != null && loadedVba.Modules[moduleName] != null;

        // The Scripting Runtime reference cannot be added via the current API,
        // so we only report its presence (expected to be false).
        bool referenceFound = false;
        if (loadedVba != null)
        {
            foreach (VbaReference reference in loadedVba.References)
            {
                if (!string.IsNullOrEmpty(reference.LibId) && reference.LibId.Contains("Scripting"))
                {
                    referenceFound = true;
                    break;
                }
            }
        }

        Console.WriteLine($"Reference added: {referenceFound}");
        Console.WriteLine($"Module added: {moduleFound}");
    }
}
