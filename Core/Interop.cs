using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace Qisope.COMServer.Core
{
	public class Interop
	{
		public const string OLE32_DLL = "ole32.dll";
		public const string USER32_DLL = "User32.dll";
		private const string SHELL32_DLL = "shell32.dll";

		public const string IID_ICLASSFACTORY = "00000001-0000-0000-C000-000000000046";
		public const string IID_IUNKNOWN = "00000000-0000-0000-C000-000000000046";
		public const string IID_IDISPATCH = "00020400-0000-0000-C000-000000000046";

		public const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
		public const int E_NOINTERFACE = unchecked((int)0x80004002);

		[DllImport(OLE32_DLL)]
		public static extern int CoRegisterClassObject(
				ref Guid rclsid,
				[MarshalAs(UnmanagedType.Interface)] IClassFactory pUnk,
				CLSCTX dwClsContext,
				REGCLS flags,
				out uint lpdwRegister);

		[DllImport(OLE32_DLL)]
		public static extern UInt32 CoRevokeClassObject(uint dwRegister);

		[DllImport(OLE32_DLL)]
		public static extern int CoResumeClassObjects();

		[DllImport(SHELL32_DLL)]
		public static extern void SHChangeNotify(HChangeNotifyEventID wEventId, HChangeNotifyFlags uFlags, IntPtr dwItem1, IntPtr dwItem2);
	}

	[Flags]
	public enum CLSCTX : uint
	{
		INPROC_SERVER = 0x1,
		INPROC_HANDLER = 0x2,
		LOCAL_SERVER = 0x4,
		INPROC_SERVER16 = 0x8,
		REMOTE_SERVER = 0x10,
		INPROC_HANDLER16 = 0x20,
		RESERVED1 = 0x40,
		RESERVED2 = 0x80,
		RESERVED3 = 0x100,
		RESERVED4 = 0x200,
		NO_CODE_DOWNLOAD = 0x400,
		RESERVED5 = 0x800,
		NO_CUSTOM_MARSHAL = 0x1000,
		ENABLE_CODE_DOWNLOAD = 0x2000,
		NO_FAILURE_LOG = 0x4000,
		DISABLE_AAA = 0x8000,
		ENABLE_AAA = 0x10000,
		FROM_DEFAULT_CONTEXT = 0x20000,
		ACTIVATE_32_BIT_SERVER = 0x40000,
		ACTIVATE_64_BIT_SERVER = 0x80000
	}

	[Flags]
	public enum REGCLS : uint
	{
		SINGLEUSE = 0x0,
		MULTIPLEUSE = 0x1,
		MULTI_SEPARATE = 0x2,
		SUSPENDED = 0x4,
		SURROGATE = 0x8,
	}

	[StructLayout(LayoutKind.Sequential)]
	public struct POINT
	{
		public int x;
		public int y;

		public POINT(int x, int y)
		{
			this.x = x;
			this.y = y;
		}

		public static explicit operator Point(POINT p)
		{
			return new Point(p.x, p.y);
		}

		public static explicit operator POINT(Point p)
		{
			return new POINT(p.X, p.Y);
		}
	}

	[Flags]
	public enum HChangeNotifyEventID
	{
		SHCNE_ALLEVENTS = 0x7FFFFFFF,
		SHCNE_ASSOCCHANGED = 0x08000000,
		SHCNE_ATTRIBUTES = 0x00000800,
		SHCNE_CREATE = 0x00000002,
		SHCNE_DELETE = 0x00000004,
		SHCNE_DRIVEADD = 0x00000100,
		SHCNE_DRIVEADDGUI = 0x00010000,
		SHCNE_DRIVEREMOVED = 0x00000080,
		SHCNE_EXTENDED_EVENT = 0x04000000,
		SHCNE_FREESPACE = 0x00040000,
		SHCNE_MEDIAINSERTED = 0x00000020,
		SHCNE_MEDIAREMOVED = 0x00000040,
		SHCNE_MKDIR = 0x00000008,
		SHCNE_NETSHARE = 0x00000200,
		SHCNE_NETUNSHARE = 0x00000400,
		SHCNE_RENAMEFOLDER = 0x00020000,
		SHCNE_RENAMEITEM = 0x00000001,
		SHCNE_RMDIR = 0x00000010,
		SHCNE_SERVERDISCONNECT = 0x00004000,
		SHCNE_UPDATEDIR = 0x00001000,
		SHCNE_UPDATEIMAGE = 0x00008000,
	}

	[Flags]
	public enum HChangeNotifyFlags
	{
		SHCNF_DWORD = 0x0003,
		SHCNF_IDLIST = 0x0000,
		SHCNF_PATHA = 0x0001,
		SHCNF_PATHW = 0x0005,
		SHCNF_PRINTERA = 0x0002,
		SHCNF_PRINTERW = 0x0006,
		SHCNF_FLUSH = 0x1000,
		SHCNF_FLUSHNOWAIT = 0x2000
	}
}