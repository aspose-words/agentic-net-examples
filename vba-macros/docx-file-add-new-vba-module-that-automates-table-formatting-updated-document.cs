using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a VBA project if the document does not already contain one.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
            doc.VbaProject.Name = "AutomationProject";
        }

        // Create a new VBA module that formats all tables in the document.
        VbaModule module = new VbaModule();
        module.Name = "TableFormatter";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub AutoFormatTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        tbl.Rows.HeightRule = wdRowHeightExactly
        tbl.Rows.Height = InchesToPoints(0.25)
        tbl.Borders.Enable = True
    Next tbl
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the updated document as a macro‑enabled file.
        doc.Save("Output.docm");
    }
}
