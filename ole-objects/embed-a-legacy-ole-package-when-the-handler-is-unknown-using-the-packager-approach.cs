using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    // OLE constants
    private const uint OLERENDER_DRAW = 0;
    private const uint COINIT_APARTMENTTHREADED = 0x2;

    // COM interfaces / CLSIDs
    private static readonly Guid IID_IOleObject = new Guid("00000112-0000-0000-C000-000000000046");

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void OleCreateFromFile(
        ref Guid rclsid,
        string lpszFileName,
        ref Guid riid,
        uint renderopt,
        IntPtr pFormatEtc,
        IntPtr pClientSite,
        out IntPtr ppvObj);

    [DllImport("ole32.dll", SetLastError = true)]
    private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

    [DllImport("ole32.dll")]
    private static extern void CoUninitialize();

    public static void Main()
    {
        // Initialize COM for the current thread
        int hr = CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
        if (hr < 0)
        {
            Console.WriteLine($"CoInitializeEx failed with HRESULT 0x{hr:X8}");
            return;
        }

        string tempFile = Path.Combine(Path.GetTempPath(), "sample.txt");
        File.WriteAllText(tempFile, "This is a sample file for OLE packaging.");

        // CLSID_NULL is just an empty GUID; we need a mutable variable to pass by ref
        Guid clsidNull = Guid.Empty;
        Guid iidOleObject = IID_IOleObject;

        IntPtr oleObjectPtr = IntPtr.Zero;
        try
        {
            OleCreateFromFile(
                ref clsidNull,
                tempFile,
                ref iidOleObject,
                OLERENDER_DRAW,
                IntPtr.Zero,
                IntPtr.Zero,
                out oleObjectPtr);

            if (oleObjectPtr != IntPtr.Zero)
            {
                // Convert the returned IUnknown pointer to a managed object
                object oleObject = Marshal.GetObjectForIUnknown(oleObjectPtr);
                Console.WriteLine("OLE object created using Packager: " + oleObject.GetType().FullName);

                // Release the COM object
                Marshal.ReleaseComObject(oleObject);
            }
            else
            {
                Console.WriteLine("Failed to create OLE object (null pointer returned).");
            }
        }
        catch (COMException ex)
        {
            Console.WriteLine("COMException: " + ex.Message);
        }
        finally
        {
            // Release the raw pointer if it was obtained
            if (oleObjectPtr != IntPtr.Zero)
            {
                Marshal.Release(oleObjectPtr);
            }

            // Clean up temporary file
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }

            // Uninitialize COM for this thread
            CoUninitialize();
        }
    }
}
