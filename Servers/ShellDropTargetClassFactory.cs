using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Qisope.COMServer.Core;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace Qisope.COMServer.Servers
{
	public class ShellDropTargetClassFactory : IClassFactory
	{
		private static Func<IDataObject, DragDropEffects> s_dragEnterHandler;
		private static Func<IDataObject, DragDropEffects> s_dragDropHandler;

		public static void RegisterHandlers(Func<IDataObject, DragDropEffects> dragEnterHandler, Func<IDataObject, DragDropEffects> dragDropHandler)
		{
			s_dragEnterHandler = dragEnterHandler;
			s_dragDropHandler = dragDropHandler;
		}

		public int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
		{
			ppvObject = IntPtr.Zero;

			if (pUnkOuter != IntPtr.Zero)
			{
				Marshal.ThrowExceptionForHR(Interop.CLASS_E_NOAGGREGATION);
			}

			if (riid == new Guid(ShellDropTarget.INTERFACE_CLSID) || riid == new Guid(Interop.IID_IDISPATCH) || riid == new Guid(Interop.IID_IUNKNOWN))
			{
				var shellDropTarget = new ShellDropTarget(s_dragEnterHandler, s_dragDropHandler);
				ppvObject = Marshal.GetComInterfaceForObject(shellDropTarget, typeof(IDropTarget));
			}
			else
			{
				Marshal.ThrowExceptionForHR(Interop.E_NOINTERFACE);
			}

			return 0;
		}

		public int LockServer(bool fLock)
		{
			return 0;
		}
	}
}