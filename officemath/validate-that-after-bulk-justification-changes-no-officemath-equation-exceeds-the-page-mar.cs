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

        // Insert several simple equations using the deterministic EQ-field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)"); // Fraction 1/2
        InsertEquation(builder, @"\r(3,x)"); // Cube root of x
        InsertEquation(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Ensure each equation is displayed on its own line.
        foreach (OfficeMath om in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Display;
                // Set a justification that will be applied to all equations.
                om.Justification = OfficeMathJustification.CenterGroup;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathJustification.docx");
        doc.Save(outputPath);

        // Reload the document to ensure layout is up‑to‑date.
        Document loadedDoc = new Document(outputPath);
        loadedDoc.UpdatePageLayout();

        // Determine the maximum allowed width for an equation (page width minus margins).
        Section section = loadedDoc.FirstSection;
        double pageWidth = section.PageSetup.PageWidth; // in points
        double maxEquationWidth = pageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;

        // Validate that no top‑level OfficeMath exceeds the margin limits.
        foreach (OfficeMath om in loadedDoc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType != MathObjectType.OMathPara)
                continue; // Skip nested math objects.

            OfficeMathRenderer renderer = new OfficeMathRenderer(om);
            double equationWidth = renderer.SizeInPoints.Width;

            if (equationWidth > maxEquationWidth + 0.1) // small tolerance
            {
                throw new InvalidOperationException(
                    $"Equation exceeds page margins. Width: {equationWidth} pts, Max allowed: {maxEquationWidth} pts.");
            }
        }

        // If we reach this point, all equations fit within the margins.
        Console.WriteLine("All OfficeMath equations are within page margin limits.");
    }

    // Helper method that inserts an EQ field, converts it to OfficeMath, and removes the field.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);
        // Return the builder to the paragraph after the field.
        builder.MoveTo(field.Start.ParentNode);
        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }
        // Add a new paragraph after the equation for readability.
        builder.InsertParagraph();
    }
}
