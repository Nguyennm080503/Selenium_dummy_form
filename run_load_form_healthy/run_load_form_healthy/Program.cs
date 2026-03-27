using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        string excelFile = "load_data.xlsx";
        var workbook = new XLWorkbook(excelFile);
        var worksheet = workbook.Worksheet(1);

        IWebDriver driver = new ChromeDriver();
        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

        driver.Navigate().GoToUrl("https://dgnc.hcdc.vn/alcohol");

        int row = 4632;
        int count = 1;
        List<int> countError = new List<int>();
        while (row < 5217)
        {
            var name = worksheet.Cell(row, 1).GetString();
            var birthday_text = worksheet.Cell(row, 2).GetString();
            var address = worksheet.Cell(row, 3).GetString();
            var phone = worksheet.Cell(row, 4).GetString();

            //if (string.IsNullOrWhiteSpace(name) &&
            //    string.IsNullOrWhiteSpace(birthday_text) &&
            //    string.IsNullOrWhiteSpace(address) &&
            //    string.IsNullOrWhiteSpace(phone))
            //{
            //    Console.WriteLine("Gặp dòng trống => Dừng đọc Excel.");
            //    break;
            //}


            try
            {

                var (day, month, year) = ParseBirthday(birthday_text);
                if (string.IsNullOrEmpty(day) || string.IsNullOrEmpty(month) || string.IsNullOrEmpty(year))
                {
                    Console.WriteLine($"Ngày sinh không hợp lệ: {birthday_text}");
                    countError.Add(row);
                    row++;
                    continue;
                }

                string gender = DetectGender(name);

                // Điền họ tên
                var nameBox = FindElementSafe(wait, By.Name("fullName"));
                if (nameBox != null)
                {
                    nameBox.Clear();
                    nameBox.SendKeys(NameGenerate(name));
                }

                // Chọn ngày/tháng/năm
                SelectCustomDropdown(wait, driver, ":R5dmmmja:-form-item", day);
                SelectCustomDropdown(wait, driver, ":R9dmmmja:-form-item", month);
                SelectCustomOtherDropdownById(wait, driver, ":Rddmmmja:-form-item", year);

                // Giới tính
                SelectGender(wait, driver, gender);

                // Điền số điện thoại
                var phoneBox = FindElementSafe(wait, By.Name("phone"));
                if (phoneBox != null)
                {
                    phoneBox.Clear();
                    phoneBox.SendKeys(PhoneGenerate(phone));
                }

                // Địa chỉ
                SelectCustomOtherDropdownById(wait, driver, ":R2tmmmja:-form-item", "phường thủ đức");

                var addressBox = FindElementSafe(wait, By.Name("address"));
                if (addressBox != null)
                {
                    addressBox.Clear();
                    addressBox.SendKeys(address);
                }

                // Submit step 1
                ClickContinue(wait, "//button[@type='submit' and contains(text(),'Tiếp tục')]", driver, "lần 1", row);

                // Option 1
                SelectOption(driver, "//button[@role='radio' and @value='989fe35c-e649-493f-8468-dbf4a08c7ba4']", wait);
                ClickContinue(wait, "//button[@type='button' and contains(text(),'Tiếp tục')]", driver, "lần 2", row);

                // Option 2
                SelectOption(driver, "//button[@role='radio' and @value='e3368e94-84dc-471f-9891-7312202e2926']", wait);
                ClickContinue(wait, "//button[@type='button' and contains(text(),'Tiếp tục')]", driver, "lần 3", row);

                // Option 3
                SelectOption(driver, "//button[@role='radio' and @value='5a97d5f0-207e-4baa-9d2e-37b0a4315f89']", wait);

                // Gửi form
                ClickContinue(wait, "//button[@type='submit' and contains(text(),'Gửi')]", driver, "Gửi", row);
            }
            catch (Exception ex)
            {
                countError.Add(row);
                Console.WriteLine($"Lỗi ở dòng {row}: {ex.Message}");
            }
            finally
            {
                count++;
                row++;
                driver.Navigate().Refresh();
                Thread.Sleep(1000);
            }
        }
        Console.WriteLine($"Date run : {DateTime.Now}");
        Console.WriteLine($"Success line : {count}");
        Console.WriteLine($"Error line : {string.Join(", ", countError)}");
        driver.Quit();
    }

    // ==================== Helper ====================
    static IWebElement FindElementSafe(WebDriverWait wait, By by)
    {
        try
        {
            return wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
        }
        catch
        {
            return null;
        }
    }

    static (string day, string month, string year) ParseBirthday(string birthday_text)
    {
        if (string.IsNullOrWhiteSpace(birthday_text)) return ("1", "1", "2000");
        string[] parts = birthday_text.Split(' ')[0].Split('/');
        if (parts.Length < 3) return ("1", "1", birthday_text);
        return (int.Parse(parts[1]).ToString(), int.Parse(parts[0]).ToString(), parts[2]);
    }

    static void SelectCustomDropdown(WebDriverWait wait, IWebDriver driver, string elementId, string optionText)
    {
        try
        {
            var button = FindElementSafe(wait, By.Id(elementId));
            var select = button.FindElement(By.XPath("./following-sibling::select"));
            var selectElement = new SelectElement(select);
            selectElement.SelectByText(optionText);
        }
        catch { }
    }

    static void SelectCustomOtherDropdownById(WebDriverWait wait, IWebDriver driver, string elementId, string optionValue)
    {
        try
        {
            var dropdown = FindElementSafe(wait, By.Id(elementId));
            dropdown.Click();
            Thread.Sleep(500);
            var option = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath($"//div[@role='option' and translate(@data-value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')='{optionValue.ToLower()}']")));
            option.Click();
        }
        catch { }
    }

    static void SelectGender(WebDriverWait wait, IWebDriver driver, string gender)
    {
        try
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string xpath = gender.ToLower().Contains("nữ") || gender.ToLower().Contains("nu")
                ? "//button[@role='radio' and @value='1']"
                : "//button[@role='radio' and @value='0']";

            var btn = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
            js.ExecuteScript("arguments[0].click();", btn);
        }
        catch { }
    }

    static void SelectOption(IWebDriver driver, string xpath, WebDriverWait wait)
    {
        var option = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", option);
        Thread.Sleep(300);
    }

    static void ClickContinue(WebDriverWait wait, string xpath, IWebDriver driver, string stepName, int? row)
    {
        try
        {
            var button = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(xpath)));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", button);
            Thread.Sleep(500);

            if (stepName == "Gửi")
            {
                var resultElement = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("p.leading-7")));
                Console.WriteLine($"Kết quả: {resultElement.Text}, Dòng {row}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Không thể click {stepName}: {ex.Message}");
        }
    }

    static string DetectGender(string fullName)
    {
        fullName = fullName.ToLower();
        string[] maleKeywords = { "văn", "hữu", "quốc", "minh", "đức" };
        string[] femaleKeywords = { "thị", "ngọc", "mai", "anh", "hoa" };
        foreach (var k in maleKeywords) if (fullName.Contains(k)) return "Nam";
        foreach (var k in femaleKeywords) if (fullName.Contains(k)) return "Nữ";
        return "Nam";
    }

    static string PhoneGenerate(string phone) => phone.Length == 9 ? "0" + phone : "0123456789";
    static string NameGenerate(string name) => string.IsNullOrEmpty(name) ? "Nguyễn Văn An" : name;
}
