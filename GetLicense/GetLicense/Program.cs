using System;
using Microsoft.Win32;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Text;
using System.Security.Cryptography;

class Program
{
    static void Main()
    {
        try
        {
            string pcName = Environment.MachineName;
            string osVersion = GetOSVersion();
            string windowsKey = GetWindowsProductKey();
            var officeInfo = GetOfficeInfo();
            string macAddress = GetMacAddress();

            string content = $"Інформація про систему\n" +
                           $"-------------------\n" +
                           $"Ім'я ПК: {pcName}\n" +
                           $"ОС: {osVersion}\n" +
                           $"Ключ Windows: {windowsKey}\n" +
                           $"\nІнформація про Office\n" +
                           $"-------------------\n" +
                           $"Версія: {officeInfo.Version}\n" +
                           $"Тип ліцензії: {officeInfo.LicenseType}\n" +
                           $"Ключ продукту: {officeInfo.ProductKey}\n" +
                           $"\nМережа\n" +
                           $"-------------------\n" +
                           $"MAC-адреса: {macAddress}\n" +
                           $"\nДата збору: {DateTime.Now}";

            string fileName = $"Office_System_Info_{pcName}.txt";
            File.WriteAllText(fileName, content, Encoding.UTF8);
            Console.WriteLine($"Інформацію збережено у файл: {fileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка: {ex.Message}");
        }
    }

    static string GetOSVersion()
    {
        try
        {
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
            {
                // Отримуємо всі необхідні значення з реєстру
                string productName = key?.GetValue("ProductName")?.ToString() ?? "";
                string displayVersion = key?.GetValue("DisplayVersion")?.ToString() ?? "";
                string currentBuild = key?.GetValue("CurrentBuild")?.ToString() ?? "";
                int ubr = 0;
                int.TryParse(key?.GetValue("UBR")?.ToString(), out ubr); // Оновлення Build Revision

                // Визначаємо, чи це Windows 11
                if (productName.Contains("Windows 10") &&
                    int.TryParse(currentBuild, out int buildNumber) &&
                    buildNumber >= 22000)
                {
                    productName = productName.Replace("Windows 10", "Windows 11");
                }

                // Форматуємо вихідний рядок
                string versionInfo = productName;

                if (!string.IsNullOrEmpty(displayVersion))
                    versionInfo += $" ({displayVersion})";

                versionInfo += $" [Build {currentBuild}";

                if (ubr > 0)
                    versionInfo += $".{ubr}";

                versionInfo += "]";

                return versionInfo;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Помилка при отриманні версії ОС: {ex.Message}");
        }

        // Резервний спосіб через WMI
        try
        {
            using (var searcher = new ManagementObjectSearcher("SELECT Caption, Version FROM Win32_OperatingSystem"))
            {
                foreach (ManagementObject os in searcher.Get())
                {
                    string caption = os["Caption"].ToString();
                    string version = os["Version"].ToString();

                    if (caption.Contains("Windows 10") && version.StartsWith("10.0.22"))
                    {
                        caption = caption.Replace("Windows 10", "Windows 11");
                    }

                    return $"{caption} (Build {version})";
                }
            }
        }
        catch { }

        return "Не вдалося визначити версію ОС";
    }

    static string GetWindowsProductKey()
    {
        try
        {
            // Спосіб 1: Для цифрових ліцензій та OEM
            string keyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(keyPath))
            {
                byte[] digitalProductId = key.GetValue("DigitalProductId") as byte[];
                if (digitalProductId != null && digitalProductId.Length >= 15)
                {
                    string decodedKey = DecodeWindowsKey(digitalProductId);
                    if (!string.IsNullOrEmpty(decodedKey))
                        return decodedKey;
                }
            }

            // Спосіб 2: Для старих версій Windows
            using (var searcher = new ManagementObjectSearcher("SELECT OA3xOriginalProductKey FROM SoftwareLicensingService"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    if (obj["OA3xOriginalProductKey"] != null)
                        return obj["OA3xOriginalProductKey"].ToString();
                }
            }
        }
        catch { }
        return "Не вдалося отримати ключ";
    }

    static (string Version, string LicenseType, string ProductKey) GetOfficeInfo()
    {
        try
        {
            // Спершу перевіряємо Click-to-Run (Office 365)
            var c2rInfo = GetClickToRunInfo();
            if (!string.IsNullOrEmpty(c2rInfo.Version))
                return c2rInfo;

            // Потім перевіряємо традиційні інсталяції
            return GetTraditionalOfficeInfo();
        }
        catch
        {
            return ("Не вдалося визначити", "", "");
        }
    }

    static (string Version, string LicenseType, string ProductKey) GetClickToRunInfo()
    {
        try
        {
            using (RegistryKey c2rKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
            {
                if (c2rKey != null)
                {
                    string productLine = c2rKey.GetValue("ProductLine")?.ToString() ?? "Office 365";
                    string version = c2rKey.GetValue("VersionToReport")?.ToString() ?? "Не вдалося визначити";
                    return ($"{productLine} (Click-to-Run)", "Цифрова ліцензія", "Ключ зберігається в обліковому записі Microsoft");
                }
            }
        }
        catch { }
        return ("", "", "");
    }

    static (string Version, string LicenseType, string ProductKey) GetTraditionalOfficeInfo()
    {
        try
        {
            string[] officeVersions = { "16.0", "15.0", "14.0", "12.0" };
            foreach (string version in officeVersions)
            {
                using (RegistryKey baseKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Microsoft\Office\{version}\Registration"))
                {
                    if (baseKey != null)
                    {
                        foreach (string subKeyName in baseKey.GetSubKeyNames())
                        {
                            using (RegistryKey subKey = baseKey.OpenSubKey(subKeyName))
                            {
                                string productName = subKey?.GetValue("ProductName")?.ToString() ?? "Microsoft Office";
                                string licenseType = subKey?.GetValue("LicenseType")?.ToString() ?? "Невідомо";
                                byte[] digitalProductId = subKey?.GetValue("DigitalProductId") as byte[];

                                string productKey = "Не вдалося отримати";
                                if (digitalProductId != null && digitalProductId.Length > 0)
                                {
                                    productKey = DecodeOfficeKey(digitalProductId);
                                }

                                return ($"{productName} (версія {version})", licenseType, productKey);
                            }
                        }
                    }
                }
            }
        }
        catch { }
        return ("Microsoft Office не знайдено", "", "");
    }

    static string DecodeWindowsKey(byte[] digitalProductId)
    {
        const string keyChars = "BCDFGHJKMPQRTVWXY2346789";
        char[] productKey = new char[29];

        try
        {
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
        catch
        {
            return "Помилка декодування ключа";
        }
    }

    static string DecodeOfficeKey(byte[] digitalProductId)
    {
        const string keyChars = "BCDFGHJKMPQRTVWXY2346789";
        char[] productKey = new char[25];

        try
        {
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

            return new string(productKey).Insert(5, "-").Insert(11, "-").Insert(17, "-").Insert(23, "-");
        }
        catch
        {
            return "Помилка декодування ключа";
        }
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
                    if (!string.IsNullOrEmpty(macAddress) && macAddress.Length >= 12)
                    {
                        StringBuilder formattedMac = new StringBuilder();
                        for (int i = 0; i < macAddress.Length; i += 2)
                        {
                            if (i + 2 <= macAddress.Length)
                            {
                                formattedMac.Append(macAddress.Substring(i, 2));
                                if (i + 2 < macAddress.Length)
                                    formattedMac.Append("-");
                            }
                        }
                        return formattedMac.ToString();
                    }
                }
            }
        }
        catch { }
        return "Не вдалося отримати";
    }
}