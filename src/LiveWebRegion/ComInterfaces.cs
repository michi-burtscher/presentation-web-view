using System;
using System.Runtime.InteropServices;

namespace LiveWebRegion
{
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }

    // IDTExtensibility2 (dual, DispId 1..5). Interface pointers are taken as IntPtr:
    // letting the interop stub auto-marshal them to typed/object params crashes the
    // CLR while hosted in PowerPoint (0x80131506). We wrap them ourselves instead.
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [TypeLibType(TypeLibTypeFlags.FHidden | TypeLibTypeFlags.FDual | TypeLibTypeFlags.FDispatchable)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(IntPtr Application, ext_ConnectMode ConnectMode, IntPtr AddInInst, IntPtr custom);

        [DispId(2)]
        void OnDisconnection(ext_DisconnectMode RemoveMode, IntPtr custom);

        [DispId(3)]
        void OnAddInsUpdate(IntPtr custom);

        [DispId(4)]
        void OnStartupComplete(IntPtr custom);

        [DispId(5)]
        void OnBeginShutdown(IntPtr custom);
    }

    // Office.IRibbonExtensibility — supplies the Ribbon UI XML. Only string in/out,
    // so it marshals safely.
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }
}
