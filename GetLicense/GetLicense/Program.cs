using System;
using System.Management;
using Microsoft.Win32;
using System.IO;
using System.Net.NetworkInformation;
using System.Linq;

class Program
{
    static void Main()
    {
        string pcName = Environment.MachineName;
        string windowsKey = GetRealWindowsProductKey();
        string osVersion = GetOSVersion();
        string officeVersion = GetOfficeVersion();
        string officeKey = GetOfficeProductKey();
        string macAddress = GetMacAddress();

        string content = $"Ім'я ПК: {pcName}\n" +
                        $"Версія ОС: {osVersion}\n" +
                        $"Ключ ліцензії Windows: {windowsKey}\n" +
                        $"Версія MS Office: {officeVersion}\n" +
                        $"Ключ продукту MS Office: {officeKey}\n" +
                        $"MAC-адреса: {macAddress}";

        File.WriteAllText($"system_info_{pcName}.txt", content);
        Console.WriteLine($"Інформацію збережено у файлі: system_info_{pcName}.txt");
    }

    static string GetOSVersion()
    {
        try
        {
            using (var searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem"))
            {
                foreach (ManagementObject os in searcher.Get())
                {
                    return os["Caption"].ToString();
                }
            }
            return "Не вдалося визначити версію ОС";
        }
        catch { return "Помилка отримання версії ОС"; }
    }

    static string GetOfficeVersion()
    {
        try
        {
            string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
            foreach (var version in officeVersions)
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Microsoft\Office\{version}\Common\ProductVersion");
                if (key != null)
                {
                    string productVersion = key.GetValue("LastProduct").ToString();
                    return $"Office {GetOfficeName(version)} {productVersion}";
                }
            }
            return "Office не знайдено";
        }
        catch { return "Помилка отримання версії Office"; }
    }

    static string GetOfficeName(string version)
    {
        switch (version)
        {
            case "16.0": return "2019/2021/365";
            case "15.0": return "2013";
            case "14.0": return "2010";
            case "12.0": return "2007";
            default: return version;
        }
    }

    static string GetRealWindowsProductKey()
    {
        try
        {
            string keyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(keyPath))
            {
                byte[] digitalProductId = key.GetValue("DigitalProductId") as byte[];
                if (digitalProductId != null)
                {
                    return DecodeWindowsKey(digitalProductId);
                }
            }
            return "Ключ не знайдено в реєстрі";
        }
        catch (Exception ex)
        {
            return $"Помилка: {ex.Message}";
        }
    }

    static string DecodeWindowsKey(byte[] digitalProductId)
    {
        const string keyChars = "BCDFGHJKMPQRTVWXY2346789";
        char[] productKey = new char[29];

        for (int i = 24; i >= 0; i--)
        {
            int current = 0;
            for (int j = 14; j >= 0; j--)
            {
                current = (current << 8) | digitalProductId[j];
                digitalProductId[j] = (byte)(current / 24);
                current %= 24;
            }
            productKey[i] = keyChars[current];
        }

        return new string(productKey, 0, 5) + "-" +
               new string(productKey, 5, 5) + "-" +
               new string(productKey, 10, 5) + "-" +
               new string(productKey, 15, 5) + "-" +
               new string(productKey, 20, 5);
    }

    static string GetOfficeProductKey()
    {
        try
        {
            string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
            foreach (var version in officeVersions)
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey(
                    $@"SOFTWARE\Microsoft\Office\{version}\Registration");
                if (key != null)
                {
                    foreach (string subKeyName in key.GetSubKeyNames())
                    {
                        RegistryKey subKey = key.OpenSubKey(subKeyName);
                        byte[] digitalProductId = subKey?.GetValue("DigitalProductId") as byte[];
                        if (digitalProductId != null)
                        {
                            return DecodeOfficeKey(digitalProductId);
                        }
                    }
                }
            }
            return "Не знайдено";
        }
        catch { return "Не вдалося отримати ключ"; }
    }

    static string DecodeOfficeKey(byte[] digitalProductId)
    {
        const string keyChars = "BCDFGHJKMPQRTVWXY2346789";
        char[] productKey = new char[25];
        
        for (int i = 24; i >= 0; i--)
        {
            int current = 0;
            for (int j = 14; j >= 0; j--)
            {
                current = (current << 8) ^ digitalProductId[j];
                digitalProductId[j] = (byte)(current / 24);
                current %= 24;
            }
            productKey[i] = keyChars[current];
        }
        
        return new string(productKey).Insert(5, "-").Insert(11, "-").Insert(17, "-").Insert(23, "-");
    }

    static string GetMacAddress()
    {
        try
        {
            NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface adapter in interfaces)
            {
                if (adapter.OperationalStatus == OperationalStatus.Up && 
                    !adapter.Description.Contains("Virtual") &&
                    !adapter.Description.Contains("Pseudo"))
                {
                    string macAddress = adapter.GetPhysicalAddress().ToString();
                    if (!string.IsNullOrEmpty(macAddress))
                    {
                        return FormatMacAddress(macAddress);
                    }
                }
            }
            return "Не знайдено";
        }
        catch { return "Помилка отримання"; }
    }

    static string FormatMacAddress(string macAddress)
    {
        macAddress = macAddress.Replace("-", "").Replace(":", "");
        return string.Join("-", Enumerable.Range(0, 6)
            .Select(i => macAddress.Substring(i * 2, 2)));
    }
}