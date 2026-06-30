using System;
using System.Runtime.InteropServices;

namespace LiveWebRegion
{
    // --- COM connection-point plumbing (hand-declared, no typed PIA) ---

    [ComImport]
    [Guid("B196B284-BAB4-101A-B69C-00AA00341D07")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IConnectionPointContainer
    {
        void EnumConnectionPoints(out IntPtr ppEnum);
        void FindConnectionPoint(ref Guid riid, out IConnectionPoint ppCP);
    }

    [ComImport]
    [Guid("B196B286-BAB4-101A-B69C-00AA00341D07")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IConnectionPoint
    {
        void GetConnectionInterface(out Guid iid);
        void GetConnectionPointContainer(out IConnectionPointContainer ppCPC);
        void Advise([MarshalAs(UnmanagedType.IUnknown)] object sink, out int cookie);
        void Unadvise(int cookie);
        void EnumConnections(out IntPtr ppEnum);
    }

    // PowerPoint EApplication event dispinterface (only the member we use).
    [ComImport]
    [Guid("914934C2-5A91-11CF-8700-00AA0060263B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IEApplicationEvents
    {
        [DispId(2003)]
        void WindowBeforeDoubleClick(
            [MarshalAs(UnmanagedType.IDispatch)] object Sel,
            [MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel);
    }

    /// <summary>Sink that forwards WindowBeforeDoubleClick to a handler.</summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    internal sealed class AppEventSink : IEApplicationEvents
    {
        // Handler returns true if it handled the double-click (cancel default action).
        public Func<object, bool> OnDoubleClick;

        public void WindowBeforeDoubleClick(object Sel, ref bool Cancel)
        {
            try { if (OnDoubleClick != null && OnDoubleClick(Sel)) Cancel = true; }
            catch (Exception ex) { Log.Error("WindowBeforeDoubleClick failed", ex); }
        }
    }

    /// <summary>Advises/unadvises the EApplication connection point.</summary>
    internal sealed class AppEventConnector
    {
        private IConnectionPoint _cp;
        private int _cookie;
        private AppEventSink _sink;

        public void Connect(object app, Func<object, bool> onDoubleClick)
        {
            try
            {
                var cpc = app as IConnectionPointContainer;
                if (cpc == null) { Log.Error("App is not an IConnectionPointContainer."); return; }

                Guid iid = typeof(IEApplicationEvents).GUID;
                cpc.FindConnectionPoint(ref iid, out _cp);
                if (_cp == null) { Log.Error("EApplication connection point not found."); return; }

                _sink = new AppEventSink { OnDoubleClick = onDoubleClick };
                _cp.Advise(_sink, out _cookie);
                Log.Info("App events connected (cookie=" + _cookie + ").");
            }
            catch (Exception ex) { Log.Error("AppEventConnector.Connect failed", ex); }
        }

        public void Disconnect()
        {
            try { if (_cp != null && _cookie != 0) _cp.Unadvise(_cookie); }
            catch { }
            _cp = null; _cookie = 0; _sink = null;
        }
    }
}
