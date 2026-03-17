using System;
using Aspose.Words;
using Aspose.Words.Vba;

class AddVbaLogging
{
    static void Main()
    {
        // Load the macro-enabled document.
        Document doc = new Document("Input.docm");

        // Ensure the document actually contains VBA macros.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Standardized logging routine to be inserted at the beginning of each module.
        const string loggingRoutine = @"
Sub LogError(ByVal errMsg As String)
    ' Simple error logger that writes to a text file.
    Dim fso As Object
    Set fso = CreateObject(""Scripting.FileSystemObject"")
    Dim ts As Object
    Set ts = fso.OpenTextFile(""C:\VbaErrorLog.txt"", 8, True)
    ts.WriteLine Now & "": "" & errMsg
    ts.Close
End Sub
";

        // Iterate through all VBA modules in the project.
        VbaModuleCollection modules = doc.VbaProject.Modules;
        foreach (VbaModule module in modules)
        {
            // Retrieve the existing source code.
            string originalCode = module.SourceCode ?? string.Empty;

            // Prepend the logging routine (ensure a line break between routine and original code).
            string newCode = loggingRoutine.TrimEnd() + "\r\n\r\n" + originalCode;

            // Update the module's source code.
            module.SourceCode = newCode;
        }

        // Save the modified document.
        doc.Save("Output_WithLogging.docm");
    }
}
