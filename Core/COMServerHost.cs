using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Qisope.COMServer.Core
{
	public class COMServerHost
	{
		#region Static Instance

		private static COMServerHost s_instance;

		public static COMServerHost Instance
		{
			get { return s_instance ?? (s_instance = new COMServerHost()); }
		}

		#endregion

		private Type[] _comServerTypes;
		private Thread _runServerThread;
		private List<uint> _registrationCookies;
		private ApplicationContext _applicationContext;

		public void Run(params Type[] comServerTypes)
		{
			if (comServerTypes == null || comServerTypes.Length == 0)
			{
				return;
			}

			_comServerTypes = comServerTypes;
			_registrationCookies = new List<uint>();
			_applicationContext = new ApplicationContext();

			_runServerThread = new Thread(RunWorker) { Name = "COMServerHost", IsBackground = true };
			_runServerThread.SetApartmentState(ApartmentState.STA);
			_runServerThread.Start();
		}

		public void Stop()
		{
			var applicationContext = _applicationContext;

			if (applicationContext != null)
			{
				applicationContext.ExitThread();
			}

			GC.SuppressFinalize(this);
		}

		private void RunWorker()
		{
			try
			{
				RegisterCOMClasses();
				Application.Run(_applicationContext);
			}
			finally
			{
				UnregisterCOMClasses();
				_runServerThread = null;
				_applicationContext = null;
			}
		}

		private void RegisterCOMClasses()
		{
			foreach (var type in _comServerTypes)
			{
				RegisterCOMClass(type);
			}

			// We registered all the classes as 'suspended', now we resume them
			int hResult = Interop.CoResumeClassObjects();

			if (hResult != 0)
			{
				throw new Exception("CoResumeClassObjects failed", Marshal.GetExceptionForHR(hResult));
			}
		}

		private void RegisterCOMClass(Type type)
		{
			var registrationData = COMServerRegistrationData.ExtractFrom(type);

			if (registrationData != null)
			{
				var clsid = registrationData.CLSID;
				var classFactory = (IClassFactory)Activator.CreateInstance(registrationData.ClassFactoryType);
				uint cookie;
				int hResult = Interop.CoRegisterClassObject(ref clsid, classFactory, CLSCTX.LOCAL_SERVER, REGCLS.MULTIPLEUSE | REGCLS.SUSPENDED, out cookie);

				if (hResult != 0)
				{
					throw new Exception(string.Format("CoRegisterClassObject failed: {0}", type.Name), Marshal.GetExceptionForHR(hResult));
				}
					
				_registrationCookies.Add(cookie);
			}
		}

		private void UnregisterCOMClasses()
		{
			foreach (uint cookie in _registrationCookies)
			{
				Interop.CoRevokeClassObject(cookie);
			}

			_registrationCookies.Clear();
		}

		~COMServerHost()
		{
			Stop();
		}
	}
}