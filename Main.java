import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String filePath = "student.xlsx";
        List<Student> students = readStudentsFromExcel(filePath);
        displayScholarshipInfo(students);
    }
    private static List<Student> readStudentsFromExcel(String filePath) {
        List<Student> students = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String name = row.getCell(0).getStringCellValue();
                double currentScholarship = row.getCell(1).getNumericCellValue();
                double newScholarship = row.getCell(2).getNumericCellValue();
                students.add(new Student(name, currentScholarship, newScholarship));
            }

        } catch (IOException e) {
            System.out.println("Ошибка чтения файла: " + e.getMessage());
        }
        return students;
    }
    private static void displayScholarshipInfo(List<Student> students) {
        for (Student student : students) {
            System.out.println("Имя: " + student.getName());
            System.out.println("Текущая стипендия: " + student.getCurrentScholarship());
            System.out.println("Новая стипендия: " + student.getNewScholarship());
            System.out.println("Увеличение стипендии: " + student.getScholarshipIncrease());
        }
    }
}
