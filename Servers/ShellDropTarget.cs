using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Qisope.COMServer.Core;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace Qisope.COMServer.Servers
{
	[ComVisible(true), Guid(CLSID), COMServer(typeof (ShellDropTargetClassFactory))]
	public class ShellDropTarget : IDropTarget
	{
		public const string CLSID = "0C2D50C5-E4F9-4662-A085-FA33181BF6B0";
		public const string INTERFACE_CLSID = "00000122-0000-0000-C000-000000000046";

		private readonly Func<IDataObject, DragDropEffects> _dragDropHandler;
		private readonly Func<IDataObject, DragDropEffects> _dragEnterHandler;

		public ShellDropTarget(Func<IDataObject, DragDropEffects> dragEnterHandler, Func<IDataObject, DragDropEffects> dragDropHandler)
		{
			_dragEnterHandler = dragEnterHandler;
			_dragDropHandler = dragDropHandler;
		}

		#region IDropTarget Members

		public int DragEnter(IDataObject pDataObj, uint grfKeyState, POINT pt, ref uint pdwEffect)
		{
			pdwEffect = (uint) (_dragEnterHandler == null ? DragDropEffects.None : _dragEnterHandler(pDataObj));
			return 0;
		}

		public int Drop(IDataObject pDataObj, uint grfKeyState, POINT pt, ref uint pdwEffect)
		{
			pdwEffect = (uint) (_dragDropHandler == null ? DragDropEffects.None : _dragDropHandler(pDataObj));
			return 0;
		}

		public int DragOver(uint grfKeyState, POINT pt, ref uint pdwEffect)
		{
			return 0;
		}

		public int DragLeave()
		{
			return 0;
		}

		#endregion
	}
}