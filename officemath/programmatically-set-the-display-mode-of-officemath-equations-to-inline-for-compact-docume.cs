using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper method to insert a top‑level OfficeMath equation using the EQ‑field bootstrap workflow.
        void InsertEquation(string eqSwitch)
        {
            // Insert an EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Move to the separator and write the EQ argument.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqSwitch);

            // Update the field so that the EQ argument is recognized.
            field.Update();

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

            // Insert the OfficeMath node before the field start run.
            Node fieldStart = fieldEq.Start;
            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

            // Remove the original field (including its runs).
            fieldEq.Remove();
        }

        // Insert first equation in the first paragraph.
        InsertEquation(@"\f(1,2)"); // Simple fraction.

        // Start a new paragraph and insert a second equation.
        builder.Writeln();
        InsertEquation(@"\r(3,x)"); // Simple root.

        // Set display mode of all top‑level OfficeMath equations to Inline.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Inline;
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the document to verify the display mode.
        Document loadedDoc = new Document(outputPath);
        NodeCollection loadedMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in loadedMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.DisplayType != OfficeMathDisplayType.Inline)
            {
                throw new InvalidOperationException("One or more top‑level equations are not set to Inline display mode.");
            }
        }

        // The example completes without interactive prompts.
    }
}
