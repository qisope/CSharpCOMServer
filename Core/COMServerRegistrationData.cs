using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Qisope.COMServer.Core
{
	public class COMServerRegistrationData
	{
		public Guid CLSID { get; private set; }
		public Type ClassFactoryType { get; private set; }

		public COMServerRegistrationData(Guid clsid, Type classFactoryType)
		{
			CLSID = clsid;
			ClassFactoryType = classFactoryType;
		}

		public static COMServerRegistrationData ExtractFrom(ICustomAttributeProvider type)
		{
			var customAttributes = type.GetCustomAttributes(true);

			var guidAttribute = Array.Find(customAttributes, attr => attr is GuidAttribute) as GuidAttribute;
			var comServerAttribute = Array.Find(customAttributes, attr => attr is COMServerAttribute) as COMServerAttribute;

			if (guidAttribute != null && comServerAttribute != null)
			{
				return new COMServerRegistrationData(new Guid(guidAttribute.Value), comServerAttribute.ClassFactoryType);
			}

			return null;
		}
	}
}