using System;
using System.Collections.Specialized;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;
using Qisope.COMServer.Core;
using Qisope.COMServer.Servers;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace Qisope.COMServer
{
	public class Program
	{
		private const string FILE_EXTENSION = ".qisope";
		private const string PROG_ID = "Qisope.qisopefile";

		[STAThread]
		public static void Main()
		{
			// Here we register our global handlers for the ShellDropTarget
			// The class factory will pass them to each instance it creates.
			ShellDropTargetClassFactory.RegisterHandlers(DragEnterHandler, DragDropHandler);

			// And this line is all it takes to be ready for IDropTarget clients
			COMServerHost.Instance.Run(typeof (ShellDropTarget));

			try
			{
				// Now we set up a little test data

				// We need to create a bunch of registry keys for our file extension and COM class
				AddRegistryKeysForCOMClass();

				// Explorer isn't sitting around watching every change in the registry
				// we need to tell it that we updated the file associations
				Interop.SHChangeNotify(HChangeNotifyEventID.SHCNE_ASSOCCHANGED, HChangeNotifyFlags.SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);

				// Put a bunch of files on your desktop with our extension
				CreateSampleFiles();

				Console.Out.WriteLine("Running!  Go click on the test files on your Desktop");
				Console.Out.WriteLine("Press 'q' to quit");
				Console.Out.WriteLine();

				while (Console.ReadKey(true).KeyChar != 'q')
				{
					// stop pressing random keys, it won't get you anything
				}

				Console.Out.WriteLine("Goodbye");
			}
			finally
			{
				// Clean up the mess we made
				RemoveRegistryKeysForCOMClass();
				Interop.SHChangeNotify(HChangeNotifyEventID.SHCNE_ASSOCCHANGED, HChangeNotifyFlags.SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero);
				RemoveSampleFiles();
			}
		}

		private static DragDropEffects DragEnterHandler(IDataObject dataObject)
		{
			Console.Out.WriteLine("DragEnterHandler");
			PrintFileList(dataObject);
			return DragDropEffects.Copy;
		}

		private static DragDropEffects DragDropHandler(IDataObject dataObject)
		{
			Console.Out.WriteLine("DragDropHandler");
			PrintFileList(dataObject);
			return DragDropEffects.None;
		}

		private static void PrintFileList(IDataObject dataObject)
		{
			var dataObjectHelper = new DataObject(dataObject);
			StringCollection fileDropList = dataObjectHelper.GetFileDropList();

			foreach (string filePath in fileDropList)
			{
				Console.Out.WriteLine(" - {0}", filePath);
			}

			Console.Out.WriteLine();
		}

		#region Set up test data

		private static void CreateSampleFiles()
		{
			RemoveSampleFiles();

			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

			for (int i = 0; i < 5; i++)
			{
				File.Create(Path.ChangeExtension(Path.Combine(desktopPath, "Test File " + i), FILE_EXTENSION)).Close();
			}
		}

		private static void RemoveSampleFiles()
		{
			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			string[] filePaths = Directory.GetFiles(desktopPath, "*.qisope");
			Array.ForEach(filePaths, File.Delete);
		}

		private static void AddRegistryKeysForCOMClass()
		{
			// ReSharper disable PossibleNullReferenceException
			RemoveRegistryKeysForCOMClass();

			RegistryKey classesKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes", true);

			RegistryKey qisopeKey = classesKey.CreateSubKey(FILE_EXTENSION);
			qisopeKey.SetValue(null, PROG_ID);
			qisopeKey.SetValue("Content Type", "application/vnd.qisope");

			RegistryKey qisopefileKey = classesKey.CreateSubKey(PROG_ID);
			qisopefileKey.SetValue(null, "Qisope File");
			RegistryKey shellKey = qisopefileKey.CreateSubKey("shell");
			RegistryKey openKey = shellKey.CreateSubKey("Open");
			RegistryKey dropTargetKey = openKey.CreateSubKey("DropTarget");
			dropTargetKey.SetValue("Clsid", string.Format("{{{0}}}", ShellDropTarget.CLSID));

			RegistryKey clsidKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID", true);
			RegistryKey classKey = clsidKey.CreateSubKey(string.Format("{{{0}}}", ShellDropTarget.CLSID));
			classKey.SetValue(null, "Example Shell Drop Target Handler", RegistryValueKind.String);

			RegistryKey localServer32Key = classKey.CreateSubKey("LocalServer32");
			localServer32Key.SetValue(null, GetAssemblyPath(), RegistryValueKind.String);
			// ReSharper restore PossibleNullReferenceException
		}

		private static void RemoveRegistryKeysForCOMClass()
		{
			// ReSharper disable PossibleNullReferenceException
			RegistryKey classesKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes", true);
			classesKey.DeleteSubKeyTree(FILE_EXTENSION, false);
			classesKey.DeleteSubKeyTree(PROG_ID, false);

			RegistryKey clsidKey = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID", true);
			clsidKey.DeleteSubKeyTree(ShellDropTarget.CLSID, false);
			// ReSharper restore PossibleNullReferenceException
		}

		private static string GetAssemblyPath()
		{
			return new Uri(Assembly.GetExecutingAssembly().Location).LocalPath;
		}

		#endregion
	}
}