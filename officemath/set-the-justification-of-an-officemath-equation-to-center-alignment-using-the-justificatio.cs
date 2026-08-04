using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class SetOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field that will later be converted to a real OfficeMath object.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments (a simple fraction 1/2) at the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(eqField.Start.ParentNode);

        // Update the field to ensure the field code is current, then convert it to OfficeMath.
        eqField.Update();
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        Paragraph parentParagraph = (Paragraph)eqField.Start.ParentNode;
        parentParagraph.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Verify we are working with a top‑level equation.
        if (officeMath.MathObjectType != MathObjectType.OMathPara)
            throw new InvalidOperationException("The created OfficeMath is not a top‑level equation.");

        // Set display type first, then justification (required order).
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Center;

        // Save the document.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "JustifiedEquation.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
