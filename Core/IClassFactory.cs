using System;
using System.Runtime.InteropServices;

namespace Qisope.COMServer.Core
{
	[ComImport]
    [ComVisible(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid(Interop.IID_ICLASSFACTORY)]
	public interface IClassFactory
	{
		[PreserveSig]
		int CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject);

		[PreserveSig]
		int LockServer(bool fLock);
	}
}