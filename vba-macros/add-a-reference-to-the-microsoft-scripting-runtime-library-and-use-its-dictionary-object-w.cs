using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Create a new VBA project and assign a name.
        VbaProject vbaProject = new VbaProject
        {
            Name = "AsposeDemoProject"
        };
        doc.VbaProject = vbaProject;

        // VBA macro source code that uses the Scripting.Dictionary object.
        // The Microsoft Scripting Runtime reference must be present for this code to compile in Word.
        string vbaCode = @"
Option Explicit

Sub UseDictionary()
    ' Create a new Dictionary object.
    Dim dict As New Scripting.Dictionary

    ' Add some key/value pairs.
    dict.Add ""Apple"", 1
    dict.Add ""Banana"", 2
    dict.Add ""Cherry"", 3

    ' Iterate and print the contents.
    Dim key As Variant
    For Each key In dict.Keys
        MsgBox ""Key: "" & key & "", Value: "" & dict(key)
    Next key
End Sub
";

        // Create a new procedural module and set its source code.
        VbaModule vbaModule = new VbaModule
        {
            Name = "DictionaryModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = vbaCode
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as a macro‑enabled .docm file.
        string outputPath = Path.Combine(outputDir, "DictionaryMacro.docm");
        doc.Save(outputPath, SaveFormat.Docm);

        // Simple verification that the document now contains macros.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"Has macros: {doc.HasMacros}");
    }
}
