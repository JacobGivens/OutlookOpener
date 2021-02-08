using System.Diagnostics;
using System.Threading;

namespace OutlookOpen {
	class Program {
		static void Main(string[] args) {
			Thread.Sleep(20000);
			do {
				Process[] outlookProcess = Process.GetProcessesByName("OUTLOOK");
				if (outlookProcess.Length == 0) {
					Process.Start("C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE");
					Thread.Sleep(30000);
				} else {
					Thread.Sleep(10000);
				}
			} while (true);
		}
	}
}
