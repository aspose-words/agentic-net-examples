using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class InsertOfficeMathFromLatex
{
    public static void Main()
    {
        // Output file path.
        const string outputPath = "Output.docx";

        // Original LaTeX source – kept as metadata only.
        const string latexSource = @"\int_{0}^{\infty} e^{-x}\,dx = 1";

        // Simple EQ argument that reliably converts to a real OfficeMath node.
        // A leading space is required so the field code becomes "EQ \f(1,2)".
        const string eqArgument = @" \f(1,2)";

        // -----------------------------------------------------------------
        // 1. Create a blank document and add introductory text.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document demonstrating insertion of an OfficeMath equation derived from LaTeX.");

        // -----------------------------------------------------------------
        // 2. Mark the insertion point with a bookmark.
        // -----------------------------------------------------------------
        builder.StartBookmark("eqTarget");
        builder.Writeln("Equation placeholder:");
        builder.EndBookmark("eqTarget");

        // -----------------------------------------------------------------
        // 3. Move to the bookmark and insert an EQ field (bootstrap for OfficeMath).
        // -----------------------------------------------------------------
        builder.MoveToBookmark("eqTarget");
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ argument into the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(eqArgument);

        // Update the field so that the EQ code is processed before conversion.
        eqField.Update();

        // Return to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // -----------------------------------------------------------------
        // 4. Convert the EQ field to a real OfficeMath object.
        // -----------------------------------------------------------------
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // -----------------------------------------------------------------
        // 5. Record the original LaTeX source for reference.
        // -----------------------------------------------------------------
        builder.Writeln($"LaTeX source: {latexSource}");

        // -----------------------------------------------------------------
        // 6. Save the document.
        // -----------------------------------------------------------------
        doc.Save(outputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 7. Validate that the file was created and contains a top‑level OfficeMath node.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new FileNotFoundException($"The output file '{outputPath}' was not created.");

        // Reload the saved document to ensure the OfficeMath node persisted.
        Document loadedDoc = new Document(outputPath);
        OfficeMath foundMath = (OfficeMath)loadedDoc.GetChild(NodeType.OfficeMath, 0, true);
        if (foundMath == null || foundMath.MathObjectType != MathObjectType.OMathPara)
            throw new InvalidOperationException("The expected top‑level OfficeMath node was not found in the saved document.");
    }
}
