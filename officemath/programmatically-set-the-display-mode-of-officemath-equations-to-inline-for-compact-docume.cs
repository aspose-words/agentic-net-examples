using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class SetOfficeMathDisplayInline
{
    public static void Main()
    {
        // Prepare file paths.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample_Inline.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with two equations using the safe
        //    EQ‑field bootstrap workflow.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation (fraction 1/2).
        InsertEquation(builder, @"\f(1,2)");
        builder.Writeln(); // start a new paragraph for the next equation

        // Second equation (cube root of x).
        InsertEquation(builder, @"\r(3,x)");
        builder.Writeln();

        // Save the sample document – this will be the input file for the
        // load‑modify‑save scenario.
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Reload the document and set the display mode of all top‑level
        //    OfficeMath nodes to Inline.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        var topLevelMath = loadedDoc.GetChildNodes(NodeType.OfficeMath, true)
                                    .OfType<OfficeMath>()
                                    .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in topLevelMath)
        {
            om.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the modified document.
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Validation.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        Document verifyDoc = new Document(outputPath);
        var verifyMath = verifyDoc.GetChildNodes(NodeType.OfficeMath, true)
                                 .OfType<OfficeMath>()
                                 .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in verifyMath)
        {
            if (om.DisplayType != OfficeMathDisplayType.Inline)
                throw new Exception("One or more OfficeMath equations are not set to Inline display mode.");
        }

        // All done.
    }

    // Inserts an EQ field with the specified argument string, converts it to a real
    // OfficeMath node, and replaces the field with the OfficeMath node.
    private static void InsertEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = field as FieldEQ;
        if (fieldEQ == null)
            throw new Exception("Failed to create an EQ field.");

        // Write the EQ arguments (prepend a space to ensure proper field syntax).
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(" " + eqArguments);

        // Update the field so that the internal representation is ready for conversion.
        fieldEQ.Update();

        // Convert the field to OfficeMath.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            throw new Exception("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        fieldEQ.Start.ParentNode.InsertBefore(officeMath, fieldEQ.Start);
        fieldEQ.Remove();

        // Move the builder back to the paragraph containing the newly inserted OfficeMath.
        builder.MoveTo(officeMath.ParentParagraph);
    }
}
