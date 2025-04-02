import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class Test01 {
    @Test
    void test01() throws IOException {
        System.out.println("Starting test...");

        // ตั้งค่าตำแหน่งของ ChromeDriver
        System.setProperty("webdriver.chrome.driver", "./chromedriver-win64/chromedriver.exe");
        System.out.println("ChromeDriver path set");

        // ตั้งค่า Chrome Options
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-gpu");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--start-maximized");
        System.out.println("Chrome options configured");

        // โหลดไฟล์ Excel
        String excelPath = "./exel/Sci.xlsx";
        WebDriver driver = null;
        try (FileInputStream fis = new FileInputStream(excelPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            System.out.println("Excel file loaded successfully");
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowNum = sheet.getLastRowNum();
            System.out.println("Total rows to process: " + rowNum);

            // เปิด WebDriver
            System.out.println("Initializing ChromeDriver...");
            driver = new ChromeDriver(options);
            driver.manage().window().maximize();
            driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);
            driver.manage().timeouts().pageLoadTimeout(30, TimeUnit.SECONDS);
            WebDriverWait wait = new WebDriverWait(driver, 2);
            System.out.println("ChromeDriver initialized successfully");

            for (int i = 1; i <= rowNum; i++) {
                try {
                    System.out.println("\nProcessing row " + i);
                    Row row = sheet.getRow(i);
                    if (row == null) {
                        System.out.println("Row " + i + " is empty, skipping...");
                        continue;
                    }

                    // ไปที่หน้าสมัคร
                    System.out.println("Navigating to signup page...");
                    driver.get("http://localhost/sc_shortcourses/signup");
                    System.out.println("Current URL: " + driver.getCurrentUrl());

                    // รอให้หน้าเว็บโหลดเสร็จ
                    System.out.println("Waiting for page to load...");
                    wait.until(ExpectedConditions.presenceOfElementLocated(By.id("nameTitleTha")));
                    System.out.println("Page loaded successfully");

                    // อ่านข้อมูลจาก Excel ตามโครงสร้างใหม่
                    System.out.println("Reading data from Excel...");
                    String prefix_th = getCellValue(row.getCell(0)); // prefix_th
                    String first_name_th = getCellValue(row.getCell(1)); // first_name_th
                    String last_name_th = getCellValue(row.getCell(2)); // last_name_th
                    String prefix_en = getCellValue(row.getCell(3)); // prefix_en
                    String first_name_en = getCellValue(row.getCell(4)); // first_name
                    String last_name_en = getCellValue(row.getCell(5)); // last_name
                    String birthDate = getCellValue(row.getCell(6)); // birthday
                    String birthMonth = getCellValue(row.getCell(7)); // birthmonth
                    String birthYear = getCellValue(row.getCell(8)); // birthyear
                    String idCard = getCellValue(row.getCell(9)); // id_card
                    String password = getCellValue(row.getCell(10)); // password
                    String mobile = getCellValue(row.getCell(11)); // phoneNo
                    String email = getCellValue(row.getCell(12)); // email
                    String address = getCellValue(row.getCell(13)); // homeNo
                    String province = getCellValue(row.getCell(14)); // province
                    String district = getCellValue(row.getCell(15)); // district
                    String subDistrict = getCellValue(row.getCell(16)); // subdistrict
                    String postalCode = getCellValue(row.getCell(17)); // post_code
                    System.out.println("Data read successfully");

                    // รอและกรอกข้อมูลในฟอร์ม
                    System.out.println("Filling form...");

                    // กรอกข้อมูลส่วนที่ 1 - อ่านข้อมูลจาก Excel
                    fillDropdown("nameTitleTha", prefix_th, wait, driver);
                    fillTextField("firstnameTha", first_name_th, wait, driver);
                    fillTextField("lastnameTha", last_name_th, wait, driver);
                    fillDropdown("nameTitleEng", prefix_en, wait, driver);
                    fillTextField("firstnameEng", first_name_en, wait, driver);
                    fillTextField("lastnameEng", last_name_en, wait, driver);

                    // กรอกข้อมูลวันเดือนปีเกิด
                    fillDropdown("birthDate", birthDate, wait, driver);
                    fillDropdown("birthMonth", birthMonth, wait, driver);
                    fillDropdown("birthYear", birthYear, wait, driver);

                    // กรอกข้อมูลส่วนที่ 2
                    fillTextField("idCard", idCard, wait, driver);
                    fillTextField("password", password, wait, driver);
                    fillTextField("mobile", mobile, wait, driver);
                    fillTextField("email", email, wait, driver);

                    // กรอกข้อมูลที่อยู่
                    fillTextField("address", address, wait, driver);
                    fillDropdown("province", province, wait, driver);
                    fillTextField("district", district, wait, driver);
                    fillTextField("subDistrict", subDistrict, wait, driver);
                    fillTextField("postalCode", postalCode, wait, driver);

                    // คลิก checkbox ยอมรับข้อตกลง
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("accept")));
                    WebElement acceptCheckbox = driver.findElement(By.id("accept"));
                    if (!acceptCheckbox.isSelected()) {
                        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", acceptCheckbox);
                    }

                    // คลิกปุ่มลงทะเบียน
                    System.out.println("Clicking register button...");
                    wait.until(
                            ExpectedConditions.elementToBeClickable(By.cssSelector("button.btn.btn-primary.btn-lg")));
                    WebElement submitButton = driver.findElement(By.cssSelector("button.btn.btn-primary.btn-lg"));
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", submitButton);
                    System.out.println("Registration submitted");

                    System.out.println("Form filled successfully");
                    System.out.println("Waiting 5 seconds before next iteration...");
                    TimeUnit.SECONDS.sleep(2);

                } catch (Exception e) {
                    System.out.println("Error processing row " + i + ": " + e.getMessage());
                    e.printStackTrace();
                    System.out.println("Waiting 10 seconds before retrying...");
                    TimeUnit.SECONDS.sleep(2);
                }
            }

        } catch (IOException | InterruptedException e) {
            System.out.println("Fatal error: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (driver != null) {
                System.out.println("Closing browser...");
                driver.quit();
                System.out.println("Browser closed");
            }
        }
    }

    private void fillDropdown(String elementId, String value, WebDriverWait wait, WebDriver driver) {
        try {
            wait.until(ExpectedConditions.elementToBeClickable(By.id(elementId)));
            WebElement element = driver.findElement(By.id(elementId));
            Select dropdown = new Select(element);

            // แสดงรายการตัวเลือกทั้งหมดที่มีใน dropdown
            System.out.println("Available options in " + elementId + ":");
            for (WebElement option : dropdown.getOptions()) {
                System.out.println("- " + option.getText());
            }

            // ลองเลือกตัวเลือก
            try {
                dropdown.selectByVisibleText(value);
                System.out.println("Selected " + elementId + ": " + value);
            } catch (NoSuchElementException e) {
                System.out.println("Warning: Could not find exact match for '" + value + "' in " + elementId);
                // ลองเลือกตัวเลือกแรกถ้าไม่พบตัวเลือกที่ต้องการ
                if (!dropdown.getOptions().isEmpty()) {
                    dropdown.selectByIndex(0);
                    System.out.println("Selected first option instead");
                }
            }
        } catch (Exception e) {
            System.out.println("Error selecting " + elementId + ": " + e.getMessage());
            throw e;
        }
    }

    private void fillTextField(String elementId, String value, WebDriverWait wait, WebDriver driver) {
        try {
            wait.until(ExpectedConditions.elementToBeClickable(By.id(elementId)));
            WebElement element = driver.findElement(By.id(elementId));
            element.clear();
            element.sendKeys(value);
            System.out.println("Filled " + elementId + ": " + value);
        } catch (Exception e) {
            System.out.println("Error filling " + elementId + ": " + e.getMessage());
            throw e;
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null)
            return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (cell.getNumericCellValue() == Math.floor(cell.getNumericCellValue())) {
                    // ถ้าเป็นเลขจำนวนเต็ม ให้แปลงเป็น long เพื่อรองรับเลขบัตรประชาชน
                    return String.valueOf((long) cell.getNumericCellValue());
                }
                return String.valueOf(cell.getNumericCellValue());
            default:
                return "";
        }
    }
}