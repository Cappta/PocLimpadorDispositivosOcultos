using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace PocLimpadorDispositivosOcultos
{
	class Program
	{
		private const string DEVCON_PATH = @"DevCon.exe";

		private const string ACTIVE_DEVICES_QUERY = @"resources =";
		private static readonly Regex ACTIVE_DEVICES_REGEX = new Regex(@"^([^\s]*)$");

		private const string ALL_DEVICES_QUERY = @"findall =";
		private static readonly Regex ALL_DEVICES_REGEX = new Regex(@"(.*?)\s*:.*");
		private const int ALL_DEVICES_REGEX_GROUP_INDEX = 1;

		private static readonly Regex NEW_LINE_REGEX = new Regex(@"\r\n|\n|\r");

		private static readonly string[] AVAILABLE_CLASSES = { "ports", "usb" };
		private static readonly string DEFAULT_CLASS = AVAILABLE_CLASSES[0];

		static void Main(string[] args)
		{
			var deviceClass = PromptDeviceClass();
			var stopwatch = Stopwatch.StartNew();
			var activeDevices = ObtainActiveDevicesWithRegisteredPorts(deviceClass).ToArray();
			var allDevices = ObtainAllDevicesWithRegisteredPorts(deviceClass).ToArray();
			var disconnectedDevices = allDevices.Where(device => activeDevices.Contains(device) == false).ToArray();

			var successfullyRemovedDevicesCount = default(int);
			foreach (var device in disconnectedDevices)
			{
				if (RemoveDevice(device) == true) { successfullyRemovedDevicesCount++; continue; }

				Console.WriteLine($"Erro ao remover o dispositivo \"{device}\"");
			}

			switch (successfullyRemovedDevicesCount)
			{
				case 0:
					Console.Write("Nenhum dispositivo foi removido");
					break;

				case 1:
					Console.Write("1 dispositivo desconectado foi removido");
					break;

				default:
					Console.Write($"{successfullyRemovedDevicesCount} dispositivos desconectados foram removidos");
					break;
			}
			Console.WriteLine($" em {stopwatch.ElapsedMilliseconds}ms");
			
			Console.WriteLine($"Pressione qualquer tecla para finalizar");
			Console.ReadKey();
		}

		private static string PromptDeviceClass()
		{
			var deviceClasses = string.Join(" ou ", AVAILABLE_CLASSES.Select(c => $"\"{c}\""));
			Console.WriteLine($"Digite {deviceClasses} para escolher a classe do dispositivo ou usaremos {DEFAULT_CLASS} por padrão");
			var deviceClass = Console.ReadLine().ToLower();

			if (AVAILABLE_CLASSES.Contains(deviceClass)) { return deviceClass; }
			
			if (string.IsNullOrWhiteSpace(deviceClass)) { return DEFAULT_CLASS; }
			
			Console.WriteLine($"A classe do dispositivo é inválida então vamos usar \"{DEFAULT_CLASS}\"");
			return DEFAULT_CLASS;
		}

		private static IEnumerable<string> ObtainActiveDevicesWithRegisteredPorts(string deviceClass)
		{
			var output = ExecuteDevCon(ACTIVE_DEVICES_QUERY + deviceClass);
			var matches = ACTIVE_DEVICES_REGEX.Matches(output);
			foreach (var line in NEW_LINE_REGEX.Split(output).Where(line => string.IsNullOrWhiteSpace(line) == false))
			{
				if (ACTIVE_DEVICES_REGEX.IsMatch(line) == false) { continue; }

				yield return line;
			}
		}

		private static IEnumerable<string> ObtainAllDevicesWithRegisteredPorts(string deviceClass)
		{
			var output = ExecuteDevCon(ALL_DEVICES_QUERY + deviceClass);
			var matches = ALL_DEVICES_REGEX.Matches(output);

			foreach (var match in matches.Cast<Match>())
			{
				yield return match.Groups[ALL_DEVICES_REGEX_GROUP_INDEX].Value;
			}
		}

		private static bool RemoveDevice(string device)
		{
			var output = ExecuteDevCon($"remove @\"{device}\"");

			return output.Contains(@"1 device(s) were removed.");
		}

		private static string ExecuteDevCon(string command)
		{
			var devCon = new Process();
			devCon.StartInfo.FileName = DEVCON_PATH;
			devCon.StartInfo.Arguments = command;
			devCon.StartInfo.UseShellExecute = false;
			devCon.StartInfo.RedirectStandardOutput = true;
			devCon.Start();

			return devCon.StandardOutput.ReadToEnd();
		}
	}
}
