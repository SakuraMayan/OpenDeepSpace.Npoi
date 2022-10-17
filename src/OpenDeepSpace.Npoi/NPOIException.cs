using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi
{

    /// <summary>
    /// NPOI异常
    /// </summary>
    public class NpoiException:Exception
    {
		public NpoiException()
		{

			
		}

		public NpoiException(string message) : base(message) { }


		public NpoiException(string message, Exception innerException) : base(message, innerException) { }


		public NpoiException(Exception innerException)
		{
		}
	}
}
