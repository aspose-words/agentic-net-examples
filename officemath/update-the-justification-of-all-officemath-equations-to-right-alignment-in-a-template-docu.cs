using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

public class UpdateOfficeMathJustification
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to insert a simple EQ field, convert it to real OfficeMath and keep it in the document.
        void InsertEquation(string eqArgs)
        {
            // Insert an EQ field.
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Write the arguments for the EQ field.
            builder.MoveTo(field.Separator);
            builder.Write(eqArgs);
            // Return the builder to the field start's parent (the paragraph).
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

        // Insert a few sample equations using safe EQ switches.
        InsertEquation(@"\f(1,2)");               // Simple fraction 1/2
        InsertEquation(@"\r(3,x)");               // Cube root of x
        InsertEquation(@"\i \su(n=1,5,n)");       // Integral with summation
        InsertEquation(@"\s \up8(Sup) \s \do8(Sub)"); // Superscript and subscript

        // Ensure the document is saved before further processing (optional).
        string tempPath = Path.Combine(Directory.GetCurrentDirectory(), "TempOfficeMath.docx");
        doc.Save(tempPath, SaveFormat.Docx);

        // Reload the document to simulate a typical template workflow.
        Document loadedDoc = new Document(tempPath);

        // Update justification of all top‑level OfficeMath paragraphs to Right.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Set display type to Display before changing justification.
                om.DisplayType = OfficeMathDisplayType.Display;
                om.Justification = OfficeMathJustification.Right;
            }
        }

        // Save the updated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "UpdatedOfficeMath.docx");
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The updated document was not saved correctly.");

        // Optional verification that all targeted OfficeMath nodes have Right justification.
        Document verifyDoc = new Document(outputPath);
        foreach (OfficeMath om in verifyDoc.GetChildNodes(NodeType.OfficeMath, true))
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.Justification != OfficeMathJustification.Right)
                throw new InvalidOperationException("An OfficeMath node does not have the expected Right justification.");
        }
    }
}
