using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Qisope.COMServer.Core;

namespace Qisope.COMServer.Servers
{
	[ComVisible(true), ComImport, Guid("00000122-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IDropTarget
	{
		[return: MarshalAs(UnmanagedType.I4)]
		[PreserveSig]
		int DragEnter(
			[In, MarshalAs(UnmanagedType.Interface)] IDataObject pDataObj,
			[In, MarshalAs(UnmanagedType.U4)] uint grfKeyState,
			[In, MarshalAs(UnmanagedType.Struct)] POINT pt,
			[In, Out, MarshalAs(UnmanagedType.U4)] ref uint pdwEffect);

		[return: MarshalAs(UnmanagedType.I4)]
		[PreserveSig]
		int DragOver(
			[In, MarshalAs(UnmanagedType.U4)] uint grfKeyState,
			[In, MarshalAs(UnmanagedType.Struct)] POINT pt,
			[In, Out, MarshalAs(UnmanagedType.U4)] ref uint pdwEffect);

		[return: MarshalAs(UnmanagedType.I4)]
		[PreserveSig]
		int DragLeave();

		[return: MarshalAs(UnmanagedType.I4)]
		[PreserveSig]
		int Drop(
			[In, MarshalAs(UnmanagedType.Interface)] IDataObject pDataObj,
			[In, MarshalAs(UnmanagedType.U4)] uint grfKeyState,
			[In, MarshalAs(UnmanagedType.Struct)] POINT pt,
			[In, Out, MarshalAs(UnmanagedType.U4)] ref uint pdwEffect);
	}
}