using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary macro-enabled document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "MacroDocument.docm");

        // 1. Create a blank document.
        Document doc = new Document();

        // 2. Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // 3. Add a procedural module.
        VbaModule procModule = new VbaModule();
        procModule.Name = "ProceduralModule";
        procModule.Type = VbaModuleType.ProceduralModule;
        procModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub

Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function
";
        doc.VbaProject.Modules.Add(procModule);

        // 4. Add a class module.
        VbaModule classModule = new VbaModule();
        classModule.Name = "MyClass";
        classModule.Type = VbaModuleType.ClassModule;
        classModule.SourceCode = @"
Option Explicit

Private m_Value As Long

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal v As Long)
    m_Value = v
End Property

Public Sub Increment()
    m_Value = m_Value + 1
End Sub
";
        doc.VbaProject.Modules.Add(classModule);

        // 5. Add a document module.
        VbaModule docModule = new VbaModule();
        docModule.Name = "DocumentModule";
        docModule.Type = VbaModuleType.DocumentModule;
        docModule.SourceCode = @"
Sub AutoOpen()
    MsgBox ""Document opened!""
End Sub
";
        doc.VbaProject.Modules.Add(docModule);

        // 6. Save the document in macro-enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // 7. Load the saved document.
        Document loadedDoc = new Document(docPath);

        // 8. Ensure the document contains a VBA project.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("No VBA project found in the document.");
            return;
        }

        // 9. Iterate through each VBA module and generate the report.
        VbaModuleCollection modules = loadedDoc.VbaProject.Modules;
        foreach (VbaModule module in modules)
        {
            // Guard against null source code.
            string source = module.SourceCode ?? string.Empty;

            // Normalize line endings and count lines.
            source = source.Replace("\r\n", "\n");
            int lineCount = source.Length == 0 ? 0 : source.Split('\n').Length;

            Console.WriteLine($"Module Name: {module.Name}");
            Console.WriteLine($"Module Type: {module.Type}");
            Console.WriteLine($"Lines of Code: {lineCount}");
            Console.WriteLine(new string('-', 30));
        }
    }
}
