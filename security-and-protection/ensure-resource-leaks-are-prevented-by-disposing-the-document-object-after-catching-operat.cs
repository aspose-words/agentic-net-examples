using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare temporary directory and file paths.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
        Directory.CreateDirectory(tempDir);
        string cancelledPath = Path.Combine(tempDir, "Cancelled.docx");
        string finalPath = Path.Combine(tempDir, "Final.docx");

        // Ensure any previous files are removed.
        if (File.Exists(cancelledPath)) File.Delete(cancelledPath);
        if (File.Exists(finalPath)) File.Delete(finalPath);

        // -----------------------------------------------------------------
        // First document: simulate a cancelled save operation.
        // -----------------------------------------------------------------
        Document doc = new Document();
        try
        {
            // Add simple content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This document will attempt a cancelled save.");

            // Create a cancelled token.
            CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel();

            // Aspose.Words does not have a Save overload that accepts a CancellationToken
            // in the version used for this example. To demonstrate handling of a
            // cancellation scenario, we manually throw the exception after the token is cancelled.
            if (cts.Token.IsCancellationRequested)
                throw new OperationCanceledException(cts.Token);

            // Normal save (won't be reached because of the exception above).
            doc.Save(cancelledPath, SaveOptions.CreateSaveOptions(SaveFormat.Docx));
        }
        catch (OperationCanceledException)
        {
            // Handle the cancellation. The document will go out of scope after this block,
            // ensuring any resources are released by the garbage collector.
            Console.WriteLine("Save operation was cancelled as expected.");
        }
        finally
        {
            // Explicitly release the reference; Document does not implement IDisposable.
            doc = null;
        }

        // Verify that the cancelled file was not created.
        if (File.Exists(cancelledPath))
            throw new InvalidOperationException("Cancelled file should not exist.");

        // -----------------------------------------------------------------
        // Second document: normal save to confirm proper operation.
        // -----------------------------------------------------------------
        Document finalDoc = new Document();
        DocumentBuilder finalBuilder = new DocumentBuilder(finalDoc);
        finalBuilder.Writeln("This document is saved after proper disposal handling.");
        finalDoc.Save(finalPath, SaveOptions.CreateSaveOptions(SaveFormat.Docx));

        // Validate final output exists.
        if (!File.Exists(finalPath))
            throw new FileNotFoundException("Final document was not saved correctly.");

        Console.WriteLine("Document saved successfully to: " + finalPath);
    }
}
