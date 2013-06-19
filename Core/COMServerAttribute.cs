using System;

namespace Qisope.COMServer.Core
{
	public class COMServerAttribute : Attribute
	{
		public Type ClassFactoryType { get; private set; }

		public COMServerAttribute(Type classFactoryType)
		{
			ClassFactoryType = classFactoryType;
		}
	}
}