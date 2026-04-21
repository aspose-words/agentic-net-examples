using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaMacroDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject project = new VbaProject
            {
                Name = "Aspose.Project"
            };
            doc.VbaProject = project;

            // Create a class module with a public method.
            VbaModule classModule = new VbaModule
            {
                Name = "MyClass",
                Type = VbaModuleType.ClassModule,
                SourceCode = @"
Public Sub MyMethod()
    MsgBox ""Hello from MyMethod""
End Sub
"
            };
            doc.VbaProject.Modules.Add(classModule);

            // Create a standard module that invokes the class method.
            VbaModule mainModule = new VbaModule
            {
                Name = "MainModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Public Sub RunMacro()
    Dim obj As New MyClass
    obj.MyMethod
End Sub
"
            };
            doc.VbaProject.Modules.Add(mainModule);

            // Save the document in a macro‑enabled format.
            doc.Save("VbaProject.CreateClassAndInvoke.docm");
        }
    }
}
