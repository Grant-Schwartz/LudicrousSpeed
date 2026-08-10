using System;
using System.Runtime.InteropServices;

namespace LudicrousSpeed.OutlookAddIn.Interop
{
    /// <summary>
    /// The interface every Office COM add-in has to implement, declared here
    /// rather than referenced.
    ///
    /// WHY DECLARE IT: the usual way to get this type is a reference to
    /// Extensibility / the Microsoft Add-in Designer, which is a machine-wide
    /// COM registration that build agents do not have -- and which would give
    /// us a second file to ship and resolve at load time. It is five methods
    /// with a fixed IID and fixed dispatch ids, all of which are part of
    /// Office's published contract, so restating it costs nothing and makes
    /// the add-in a single self-contained assembly.
    ///
    /// The GUID, the dual interface shape, and the dispatch ids 1-5 must match
    /// the real type library exactly, because Office calls these by id.
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object application,
            int connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            ref Array custom);

        [DispId(2)]
        void OnDisconnection(int removeMode, ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(ref Array custom);

        [DispId(4)]
        void OnStartupComplete(ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(ref Array custom);
    }
}
