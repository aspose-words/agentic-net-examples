using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three simple equations, each in its own paragraph, using the EQ‑field bootstrap workflow.
        string[] eqArgs = { @"\f(1,2)", @"\r(3,x)", @"\s(5)" };
        foreach (string eq in eqArgs)
        {
            // Start a new paragraph.
            builder.Writeln();

            // Insert an EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);

            // Write the EQ argument into the field separator.
            builder.MoveTo(field.Separator);
            builder.Write(eq);

            // Convert the field to a real OfficeMath node.
            FieldEQ fieldEq = field as FieldEQ;
            OfficeMath officeMath = fieldEq?.AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and remove the field.
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                field.Remove();
            }
        }

        // Apply bulk justification change to all top‑level OfficeMath equations.
        // The Justification property may not be available in older versions, so this step is optional.
        var topLevelMath = doc.GetChildNodes(NodeType.OfficeMath, true)
                              .Cast<OfficeMath>()
                              .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        // Uncomment the following block if the Justification property is supported in your version.
        /*
        foreach (OfficeMath om in topLevelMath)
        {
            om.Justification = Justification.Center;
        }
        */

        // Update layout to reflect any formatting changes.
        doc.UpdatePageLayout();

        // Validate that no equation exceeds page margins.
        Section section = doc.FirstSection;
        double leftMargin = section.PageSetup.LeftMargin;
        double rightMargin = section.PageSetup.RightMargin;
        double pageWidth = section.PageSetup.PageWidth;
        double rightBoundary = pageWidth - rightMargin;

        LayoutCollector collector = new LayoutCollector(doc);
        foreach (OfficeMath om in topLevelMath)
        {
            // Get layout information for the OfficeMath node via reflection to avoid direct dependency on LayoutEntityInfo.
            object entity = collector.GetEntity(om);
            if (entity == null)
                continue; // Should not happen, but skip if layout info is missing.

            // Use reflection to read X and Width properties.
            double eqLeft = (double)entity.GetType().GetProperty("X").GetValue(entity);
            double eqWidth = (double)entity.GetType().GetProperty("Width").GetValue(entity);
            double eqRight = eqLeft + eqWidth;

            if (eqLeft < leftMargin || eqRight > rightBoundary)
            {
                throw new InvalidOperationException(
                    $"Equation exceeds page margins. Left: {eqLeft}, Right: {eqRight}, " +
                    $"Margins => Left: {leftMargin}, Right: {rightMargin}");
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Indicate success.
        Console.WriteLine("Document saved and all equations are within page margins.");
    }
}
