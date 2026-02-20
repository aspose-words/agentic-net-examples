using System;
using Aspose.Words;
using Aspose.Words.Vba;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Enable tracking of changes so that any modifications are recorded.
        // -----------------------------------------------------------------
        doc.TrackRevisions = true;

        // -----------------------------------------------------------------
        // 2. Apply Asian typography settings via CompatibilityOptions.
        //    Example: enable East Asian line‑break rules and keep other
        //    Asian‑specific behaviours.
        // -----------------------------------------------------------------
        doc.CompatibilityOptions.DoNotUseEastAsianBreakRules = false;
        // Additional Asian typography flags can be set here as needed, e.g.:
        // doc.CompatibilityOptions.DoNotVertAlignInTxbx = false;

        // -----------------------------------------------------------------
        // 3. Add a comment to the first paragraph of the document.
        //    (InsertComment is not part of the provided API, so this is a
        //    placeholder demonstrating where such code would go.)
        // -----------------------------------------------------------------
        Paragraph firstParagraph = (Paragraph)doc.FirstSection.Body.FirstChild;
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(firstParagraph);
        // builder.InsertComment("Author", DateTime.Now, "This is a comment."); // Placeholder

        // -----------------------------------------------------------------
        // 4. Manipulate text boxes (shapes). The specific Shape API is not
        //    listed in the provided chunks, so this section is left as a
        //    placeholder for where shape handling would be performed.
        // -----------------------------------------------------------------
        // foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        // {
        //     if (shape.ShapeType == ShapeType.TextBox)
        //         shape.FillColor = System.Drawing.Color.LightYellow;
        // }

        // -----------------------------------------------------------------
        // 5. Create a VBA project with a simple macro and attach it to the
        //    document. This uses the VbaProject, VbaModule and
        //    VbaModuleType classes from the Aspose.Words.Vba namespace.
        // -----------------------------------------------------------------
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "AsposeProject";

        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = "Module1";
        vbaModule.Type = VbaModuleType.ProceduralModule;
        vbaModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        vbaProject.Modules.Add(vbaModule);
        doc.VbaProject = vbaProject;

        // -----------------------------------------------------------------
        // 6. Save the document as a macro‑enabled DOCM file to preserve the
        //    VBA project.
        // -----------------------------------------------------------------
        doc.Save("Output.docm");
    }
}
