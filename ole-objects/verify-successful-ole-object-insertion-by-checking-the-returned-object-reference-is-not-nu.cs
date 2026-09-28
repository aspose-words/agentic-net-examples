using System;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        object oleObject = null;
        try
        {
            Type progId = Type.GetTypeFromProgID("WScript.Shell");
            if (progId != null)
            {
                oleObject = Activator.CreateInstance(progId);
            }
        }
        catch (COMException ex)
        {
            Console.WriteLine($"COMException: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception: {ex.Message}");
        }

        if (oleObject != null)
        {
            Console.WriteLine("OLE object insertion successful: reference is not null.");
        }
        else
        {
            Console.WriteLine("Failed to insert OLE object: reference is null.");
        }

        if (oleObject != null && Marshal.IsComObject(oleObject))
        {
            Marshal.ReleaseComObject(oleObject);
        }
    }
}
