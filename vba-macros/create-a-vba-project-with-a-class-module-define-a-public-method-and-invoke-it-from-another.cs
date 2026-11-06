using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Define a class module with a public method.
        string classSource = @"
Public Sub ShowMessage()
    MsgBox ""Hello from class!""
End Sub
";

        VbaModule classModule = new VbaModule
        {
            Name = "MyClass",
            Type = VbaModuleType.ClassModule,
            SourceCode = classSource
        };
        doc.VbaProject.Modules.Add(classModule);

        // Define a standard module that invokes the class method.
        string standardSource = @"
Public Sub CallClassMethod()
    Dim obj As New MyClass
    obj.ShowMessage
End Sub
";

        VbaModule standardModule = new VbaModule
        {
            Name = "StandardModule",
            // The default Type for a newly created VbaModule is StandardModule,
            // so we omit setting the Type property to avoid compatibility issues.
            SourceCode = standardSource
        };
        doc.VbaProject.Modules.Add(standardModule);

        // Save the document as a macro‑enabled file.
        string fileName = "MacroDocument.docm";
        doc.Save(fileName, SaveFormat.Docm);

        // Reload the document to verify the VBA modules.
        Document loaded = new Document(fileName);
        Console.WriteLine($"Document saved as {fileName}");
        Console.WriteLine("VBA Modules in the document:");
        foreach (VbaModule mod in loaded.VbaProject.Modules)
        {
            Console.WriteLine($"- {mod.Name} (Type: {mod.Type})");
        }
    }
}
