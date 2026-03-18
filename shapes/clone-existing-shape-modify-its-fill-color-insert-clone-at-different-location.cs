using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert the original shape.
Shape originalShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
originalShape.Left = 100;               // Position X (points)
originalShape.Top = 100;                // Position Y (points)
originalShape.FillColor = Color.LightGray;

// Clone the shape (deep clone, include child nodes if any).
Shape clonedShape = (Shape)originalShape.Clone(true);

// Modify the fill color of the cloned shape.
clonedShape.FillColor = Color.CornflowerBlue;

// Place the cloned shape at a different location.
clonedShape.Left = originalShape.Left + 150; // Shift right.
clonedShape.Top = originalShape.Top;         // Same vertical position.

// Insert the cloned shape into the document after the original shape.
originalShape.InsertAfter(clonedShape, originalShape);

// Save the document.
doc.Save("CloneShape.docx");
