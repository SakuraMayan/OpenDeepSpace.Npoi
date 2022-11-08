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

		public string Message { get; set; }

		public int ExceptionCode { get; set; }

		public NpoiException(string message) : base(message) { Message = message; }
		public NpoiException(int code,string message) : base(message) { Message = message;ExceptionCode = code; }


		public NpoiException(string message, Exception innerException) : base(message, innerException) { Message = message; }
		public NpoiException(int code,string message, Exception innerException) : base(message, innerException) { Message = message; ExceptionCode = code; }


		public NpoiException(Exception innerException)
		{
		}
	}
}
