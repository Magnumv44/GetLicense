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

    /// <summary>
    /// Статичний метод отримання версії операційної системи
    /// </summary>
    /// <returns></returns>
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
        catch
        {
        }

        return "Не вдалося визначити версію ОС";
    }

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
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
            using (var searcher =
                   new ManagementObjectSearcher("SELECT OA3xOriginalProductKey FROM SoftwareLicensingService"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    if (obj["OA3xOriginalProductKey"] != null)
                        return obj["OA3xOriginalProductKey"].ToString();
                }
            }
        }
        catch
        {
        }

        return "Не вдалося отримати ключ";
    }

    /// <summary>
    /// Головний метод для отримання інформації про Microsoft Office
    /// Виконує послідовну перевірку: швидка перевірка на наявності → детальний аналіз на предмет звичної інсталяції ПЗ → перевірка на Click-to-Run → фінальний висновок
    /// Повертає кортеж з версією, типом ліцензії та ключем продукту
    /// </summary>
    /// <returns></returns>
    static (string Version, string LicenseType, string ProductKey) GetOfficeInfo()
    {
        // 1. Спершу перевіряємо очевидні ознаки відсутності Office
        if (!IsOfficeLikelyInstalled())
        {
            return ("Не встановлено", "", "");
        }

        // 2. Детальна перевірка на предмет звичайної інсталяції MS Office з інсталяціонного дистрибутиву
        var traditionalInfo = GetTraditionalOfficeInfoDetailed();
        if (traditionalInfo.Version != "Не встановлено")
        {
            return traditionalInfo;
        }

        // 3. Перевірка Click-to-Run
        var c2rInfo = GetClickToRunInfoStrict();
        if (c2rInfo.Version != "Не встановлено")
        {
            return c2rInfo;
        }

        // 4. Якщо жоден метод не знайшов Office
        return ("Не встановлено", "", "");
    }

    /// <summary>
    /// Швидка перевірка можливої наявності Office на системі
    /// Перевіряє ключові розділи реєстру та стандартні шляхи інсталяції
    /// Оптимізовано для швидкого виконання з мінімальним навантаженням
    /// </summary>
    /// <returns>Булеве значення результату перевірки (true or false)</returns>
    static bool IsOfficeLikelyInstalled()
    {
        try
        {
            // Швидка перевірка по ключових файлах і ключах реєстру
            string[] checkPaths =
            {
                @"SOFTWARE\Microsoft\Office",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"C:\Program Files\Microsoft Office",
                @"C:\Program Files (x86)\Microsoft Office"
            };

            foreach (var path in checkPaths)
            {
                if (path.Contains("\\"))
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(path))
                    {
                        if (key != null && key.SubKeyCount > 0) return true;
                    }
                }
                else if (Directory.Exists(path))
                {
                    return true;
                }
            }
        }
        catch
        {
        }

        return false;
    }

    /// <summary>
    /// Детальний аналіз традиційних (MSI) інсталяцій Microsoft Office
    /// Перевіряє всі версії Office від 2000 до 2021 у реєстрі
    /// Включає перевірку ProductName та ProductId для ідентифікації справжніх інсталяцій Office
    /// </summary>
    /// <returns>Строкове значення (string) з описом результату перевірки</returns>
    static (string Version, string LicenseType, string ProductKey) GetTraditionalOfficeInfoDetailed()
    {
        try
        {
            string[] officeVersions = { "16.0", "15.0", "14.0", "12.0", "11.0", "10.0", "9.0" };

            foreach (string version in officeVersions)
            {
                using (var baseKey = Registry.LocalMachine.OpenSubKey(
                           $@"SOFTWARE\Microsoft\Office\{version}\Registration"))
                {
                    if (baseKey == null) continue;

                    foreach (string subKeyName in baseKey.GetSubKeyNames().Where(x => x.StartsWith("{")))
                    {
                        using (var subKey = baseKey.OpenSubKey(subKeyName))
                        {
                            if (subKey == null) continue;

                            string productName = subKey.GetValue("ProductName")?.ToString() ?? "";
                            string productId = subKey.GetValue("ProductId")?.ToString() ?? "";

                            // Тверда перевірка, що це дійсно Office
                            if (!IsValidOfficeProduct(productName, productId))
                                continue;

                            string licenseType = subKey.GetValue("LicenseType")?.ToString() ?? "Невідомо";
                            byte[] digitalProductId = subKey.GetValue("DigitalProductId") as byte[];
                            string productKey = digitalProductId != null
                                ? DecodeOfficeKey(digitalProductId)
                                : "Не вдалося отримати";

                            return (GetOfficeVersionName(version, productName), licenseType, productKey);
                        }
                    }
                }
            }
        }
        catch
        {
        }

        return ("Не встановлено", "", "");
    }

    /// <summary>
    /// Перевірка чи є продукт справжнім Office
    /// Використовує списки допустимих продуктів і виключень
    /// Додатково перевіряє ProductId для підтвердження
    /// Запобігає помилковій ідентифікації суміжних продуктів (наприклад, Skype, Proofing Tools)
    /// </summary>
    /// <param name="productName">string</param>
    /// <param name="productId">string</param>
    /// <returns>Повертає строку типу string з назвою різновиду MS Office (Standart, Professional, Home і тд)</returns>
    static bool IsValidOfficeProduct(string productName, string productId)
    {
        if (string.IsNullOrWhiteSpace(productName)) return false;

        // Список допустимих продуктів
        string[] validProducts =
        {
            "Office", "Word", "Excel", "PowerPoint", "Outlook",
            "Access", "Publisher", "OneNote", "Visio", "Project"
        };

        // Список виключень (не Office)
        string[] excludeProducts =
        {
            "Proofing", "Proof", "Compatibility", "Converter",
            "Component", "Language", "Language Pack", "Help"
        };

        bool isOffice = validProducts.Any(p => productName.Contains(p)) &&
                        !excludeProducts.Any(p => productName.Contains(p));

        // Додаткова перевірка по ProductID
        if (!string.IsNullOrWhiteSpace(productId))
        {
            isOffice = isOffice && (productId.StartsWith("Office") ||
                                    productId.Contains("Standard") ||
                                    productId.Contains("Professional") ||
                                    productId.Contains("Home"));
        }

        return isOffice;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    static (string Version, string LicenseType, string ProductKey) GetClickToRunInfoStrict()
    {
        try
        {
            using (var c2rKey = Registry.LocalMachine.OpenSubKey(
                       @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
            {
                if (c2rKey == null) return ("Не встановлено", "", "");

                string productIds = c2rKey.GetValue("ProductReleaseIds")?.ToString() ?? "";
                string version = c2rKey.GetValue("VersionToReport")?.ToString() ?? "";

                if (!string.IsNullOrWhiteSpace(productIds) && productIds.Contains("O365"))
                {
                    return ("Microsoft 365 Apps for enterprise (Click-to-Run)",
                        "Цифрова ліцензія",
                        "Ключ зберігається в обліковому записі Microsoft");
                }
            }
        }
        catch
        {
        }

        return ("Не встановлено", "", "");
    }

    /// <summary>
    /// Метод формування версійної назви MS Office на основі позначення. Наприклад 16.0 = 2021/2019/365
    /// </summary>
    /// <param name="version">string</param>
    /// <param name="productName">string</param>
    /// <returns>Повертає строкове значення в вигляді назви версії для чіткого розуміння</returns>
    static string GetOfficeVersionName(string version, string productName)
    {
        var versionMap = new Dictionary<string, string>
        {
            ["16.0"] = "2021/2019/365",
            ["15.0"] = "2013",
            ["14.0"] = "2010",
            ["12.0"] = "2007",
            ["11.0"] = "2003",
            ["10.0"] = "2002",
            ["9.0"] = "2000"
        };

        string versionName = versionMap.TryGetValue(version, out var name) ? name : version;

        // Додаємо інформацію про LTSC
        if (productName.Contains("LTSC"))
        {
            versionName += " LTSC";
        }

        return $"{productName} ({versionName})";
    }

    /// <summary>
    /// Метод декодування ключа операційної системі, та його форматування в потрібному форматі. 
    /// </summary>
    /// <param name="digitalProductId">Масив байтів (byte array)</param>
    /// <returns>Віддає строку (string type) з ключем в вигляді "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"</returns>
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

    /// <summary>
    /// 
    /// </summary>
    /// <param name="digitalProductId"></param>
    /// <returns></returns>
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

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
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
        catch
        {
        }

        return "Не вдалося отримати";
    }
}