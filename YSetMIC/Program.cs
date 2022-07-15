using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YSetMIC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(args.Length);
            if (args.Length != 1)
            {
                Console.WriteLine("参数个数不匹配！");
                Environment.Exit(-1);
                return;
            }
            DeviceState devicestate = DeviceState.Active;
            int setState = 0x00000001;
            if (args[0] == "Active")
            {
                devicestate = DeviceState.Disabled;
            }
            else if (args[0] == "Disabled")
            {
                devicestate = DeviceState.Active;
                setState = 0x10000001;
            }
            else
            {
                Console.WriteLine($"参数错误：{args[1]}");
                Environment.Exit(-1);
            }
            IMMDeviceEnumerator mMDeviceEnumerator = new IMMDeviceEnumerator();
            IMMDeviceCollection deviceCollection;
            mMDeviceEnumerator.EnumAudioEndpoints(EDataFlow.eCapture, devicestate, out deviceCollection);
            uint count = 0;
            deviceCollection.GetCount(out count);
            for (uint i = 0; i < count; i++)
            {
                IMMDevice device;
                deviceCollection.Item(i, out device);
                string id;
                device.GetId(out id);
                var guid = id.Split('.').Last();
                string key = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{guid}";
                RegistryHelper.SetRegistryValue(Registry.LocalMachine, key, "DeviceState", setState, RegistryValueKind.DWord);
                Console.WriteLine($"{guid} Set State Success: {setState}");
            }
            Environment.Exit(0);
        }
    }
}
