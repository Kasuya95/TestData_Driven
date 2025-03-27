import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class Test01 {
    @Test
    void test01() throws IOException {
        // ตั้งค่าตำแหน่งของ ChromeDriver
        System.setProperty("webdriver.chrome.driver", "./chromedriver-win64/chromedriver.exe");

        // โหลดไฟล์ Excel
        String excelPath = "./exel/Sci.xlsx"; // ใช้ path ของไฟล์ที่อัปโหลด
        try (FileInputStream fis = new FileInputStream(excelPath);
                XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowNum = sheet.getLastRowNum(); // นับจำนวนแถวทั้งหมด

            // เปิด WebDriver เพียงตัวเดียว
            WebDriver driver = new ChromeDriver();
            driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS); // ใช้ TimeUnit สำหรับ implicitWait
            WebDriverWait wait = new WebDriverWait(driver, 10); // Selenium 3 ต้องใช้ตัวเลขเป็นวินาที

            for (int i = 1; i <= rowNum; i++) {
                Row row = sheet.getRow(i);
                if (row == null)
                    continue;

                // ไปที่หน้าสมัคร
                driver.get("http://localhost/sc_shortcourses/signup");

                // อ่านข้อมูลจาก Excel
                String nameTitleTha = getCellValue(row.getCell(1));
                String firstnameTha = getCellValue(row.getCell(2));
                String lastnameTha = getCellValue(row.getCell(3));
                String nameTitleEng = getCellValue(row.getCell(4));
                String firstnameEng = getCellValue(row.getCell(5));
                String lastnameEng = getCellValue(row.getCell(6));
                String birthDate = getCellValue(row.getCell(7));
                String birthMonth = getCellValue(row.getCell(8));
                String birthYear = getCellValue(row.getCell(9));
                String idCard = getCellValue(row.getCell(10));
                String password = getCellValue(row.getCell(11));
                String mobile = getCellValue(row.getCell(12));
                String email = getCellValue(row.getCell(13));
                String address = getCellValue(row.getCell(14));
                String province = getCellValue(row.getCell(15));
                String district = getCellValue(row.getCell(16));
                String subDistrict = getCellValue(row.getCell(17));
                String postalCode = getCellValue(row.getCell(18));

                // กรอกข้อมูลในฟอร์ม
                new Select(driver.findElement(By.id("nameTitleTha"))).selectByVisibleText(nameTitleTha);
                driver.findElement(By.id("firstnameTha")).sendKeys(firstnameTha);
                driver.findElement(By.id("lastnameTha")).sendKeys(lastnameTha);
                new Select(driver.findElement(By.id("nameTitleEng"))).selectByVisibleText(nameTitleEng);
                driver.findElement(By.id("firstnameEng")).sendKeys(firstnameEng);
                driver.findElement(By.id("lastnameEng")).sendKeys(lastnameEng);
                driver.findElement(By.id("birthDate")).sendKeys(birthDate);
                driver.findElement(By.id("birthMonth")).sendKeys(birthMonth);
                driver.findElement(By.id("birthYear")).sendKeys(birthYear);
                driver.findElement(By.id("idCard")).sendKeys(idCard);
                driver.findElement(By.id("password")).sendKeys(password);
                driver.findElement(By.id("mobile")).clear();
                driver.findElement(By.id("mobile")).sendKeys(mobile);
                driver.findElement(By.id("email")).sendKeys(email);
                driver.findElement(By.id("address")).clear();
                driver.findElement(By.id("address")).sendKeys(address);
                driver.findElement(By.id("province")).sendKeys(province);
                driver.findElement(By.id("district")).sendKeys(district);
                driver.findElement(By.id("subDistrict")).clear();
                driver.findElement(By.id("subDistrict")).sendKeys(subDistrict);
                driver.findElement(By.id("postalCode")).sendKeys(postalCode);

                // คลิก checkbox ยอมรับข้อตกลง
                ((JavascriptExecutor) driver)
                        .executeScript("document.getElementById('accept').click();");

                // รอ 5 วินาที
                TimeUnit.SECONDS.sleep(5);
            }

            // ปิด WebDriver
            driver.quit();

        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }

    // ฟังก์ชันอ่านค่าเซลล์
    private String getCellValue(Cell cell) {
        if (cell == null)
            return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            default:
                return "";
        }
    }
}
