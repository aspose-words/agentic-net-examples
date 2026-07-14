using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Rendering;

public class OfficeMathJustificationValidator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert several sample equations using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)");                     // Simple fraction 1/2
        InsertEquation(builder, @"\r(3,x)");                     // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)");            // Integral with summation
        InsertEquation(builder, @"\a \co2 \vs3 \hs3(4x,-4y,-4x,+y)"); // Array example

        // Apply bulk formatting to top‑level OfficeMath paragraphs.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in mathNodes)
        {
            if (officeMath.MathObjectType == MathObjectType.OMathPara)
            {
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                // Justification must be set after DisplayType.
                officeMath.Justification = OfficeMathJustification.CenterGroup;
            }
        }

        // Validate that no top‑level equation exceeds the printable width of the page.
        double pageWidth = doc.FirstSection.PageSetup.PageWidth;
        double leftMargin = doc.FirstSection.PageSetup.LeftMargin;
        double rightMargin = doc.FirstSection.PageSetup.RightMargin;
        double usableWidth = pageWidth - leftMargin - rightMargin;

        foreach (OfficeMath officeMath in mathNodes)
        {
            if (officeMath.MathObjectType != MathObjectType.OMathPara)
                continue; // Skip nested math objects.

            // Measure the rendered width of the equation.
            OfficeMathRenderer renderer = new OfficeMathRenderer(officeMath);
            double equationWidth = renderer.SizeInPoints.Width;

            if (equationWidth > usableWidth)
                throw new InvalidOperationException(
                    $"Equation exceeds page margins. Width: {equationWidth} pts, Usable: {usableWidth} pts.");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathJustificationResult.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
    }

    // Helper method that inserts an EQ field, converts it to a real OfficeMath node, and removes the original field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field. The field code initially contains only the "EQ" switch.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments after the field code (at the separator).
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Ensure the field is up‑to‑date before conversion.
        field.Update();

        // Return the builder to the paragraph that contains the field start.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph(); // Start a new paragraph for the next equation.

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();
    }
}
