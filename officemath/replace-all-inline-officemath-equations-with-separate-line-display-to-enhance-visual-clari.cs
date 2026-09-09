using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ReplaceInlineOfficeMath
{
    public static void Main()
    {
        // Output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample document with inline OfficeMath equations.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("This paragraph contains two inline equations:");
        builder.Font.Size = 12;
        builder.Write("The first equation is ");
        InsertInlineEquation(builder, @"\f(1,2)"); // fraction 1/2
        builder.Write(", and the second one is ");
        InsertInlineEquation(builder, @"\r(3,x)"); // cube root of x
        builder.Writeln(".");

        // Second paragraph.
        builder.Writeln("Another paragraph with an inline equation:");
        builder.Write("Euler's identity: ");
        InsertInlineEquation(builder, @"\i e^{i\pi}+1=0"); // exponential identity
        builder.Writeln(".");

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "Original.docx");
        doc.Save(originalPath);

        // 2. Load the document (simulating a separate load step).
        Document loadedDoc = new Document(originalPath);

        // 3. Change all top‑level inline OfficeMath equations to display mode.
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath officeMath in mathNodes)
        {
            if (officeMath.MathObjectType == MathObjectType.OMathPara &&
                officeMath.DisplayType == OfficeMathDisplayType.Inline)
            {
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.Left;
            }
        }

        // 4. Save the modified document.
        string resultPath = Path.Combine(outputDir, "Result.docx");
        loadedDoc.Save(resultPath);

        // 5. Simple validation – ensure the result file exists and contains at least one displayed equation.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The result document was not saved.");

        Document validationDoc = new Document(resultPath);
        NodeCollection resultMaths = validationDoc.GetChildNodes(NodeType.OfficeMath, true);
        bool hasDisplay = false;
        foreach (OfficeMath om in resultMaths)
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType == OfficeMathDisplayType.Display)
            {
                hasDisplay = true;
                break;
            }
        }

        if (!hasDisplay)
            throw new InvalidOperationException("No OfficeMath equations were set to display mode.");
    }

    // Inserts an EQ field, writes the EQ arguments, updates the field,
    // converts it to a real OfficeMath node, inserts the node, and removes the field.
    private static void InsertInlineEquation(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field (the field code initially contains only "EQ").
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments (including the leading backslash) into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that Aspose.Words parses the arguments.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);
        field.Remove();
    }
}
