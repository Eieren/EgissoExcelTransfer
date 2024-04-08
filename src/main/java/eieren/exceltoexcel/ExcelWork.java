package eieren.exceltoexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Iterator;
import javax.swing.JComboBox;
import javax.swing.JFormattedTextField;
import javax.swing.JTextArea;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWork {

    public static void validateOriginExcelFile(JTextArea console, File file) {
        try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);

            String[] expectedHeaders = {
                "СНИЛС", "Фамилия", "Имя",
                "Отчество", "Пол", "Дата рождения", "Сумма, руб."};

            for (int i = 0; i < expectedHeaders.length; i++) {
                Cell cell = firstRow.getCell(i);
                if (cell == null || !cell.getStringCellValue().equals(expectedHeaders[i])) {
                    console.append("Ошибка: Не найден пункт: " + expectedHeaders[i] + "\n");
                    return;
                }
            }

            console.append("Файл соответствует стандарту!" + "\n");
        } catch (Exception e) {
        }
    }

    public static void validatePatternExcelFile(JTextArea console, File file) {
        try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);

            String[] expectedHeaders = {
                "RecType", "assignmentFactUuid", "LMSZID",
                "categoryID", "ONMSZCode", "LMSZProviderCode", "providerCode",
                "SNILS_recip", "FamilyName_recip", "Name_recip",
                "Patronymic_recip", "Gender_recip", "BirthDate_recip", "doctype_recip",
                "doc_Series_recip", "doc_Number_recip",
                "doc_IssueDate_recip", "doc_Issuer_recip", "SNILS_reason",
                "FamilyName_reason", "Name_reason", "Patronymic_reason",
                "Gender_reason", "BirthDate_reason", "kinshipTypeCode", "doctype_reason",
                "doc_Series_reason", "doc_Number_reason", "doc_IssueDate_reason", "doc_Issuer_reason",
                "decision_date", "dateStart", "dateFinish", "usingSign",
                "criteria", "criteriaCode",
                "FormCode", "amount", "measuryCode", "monetization",
                "content", "comment", "equivalentAmount"};

            for (int i = 0; i < expectedHeaders.length; i++) {
                Cell cell = firstRow.getCell(i);
                if (cell == null || !cell.getStringCellValue().equals(expectedHeaders[i])) {
                    console.append("Ошибка: Не найден пункт: " + expectedHeaders[i] + "\n");
                    return;
                }
            }

            console.append("Файл соответствует стандарту!" + "\n");
        } catch (Exception e) {
        }
    }

    public static void dataCollectorForExcel(JComboBox recType, JComboBox lmszid,
            JComboBox categoryID, JComboBox onmszCode, JComboBox usingSign, JComboBox monetization,
            JFormattedTextField decisiondate, JFormattedTextField dateStart,
            JFormattedTextField dateFinish, String[] staticData, JTextArea console) {

        if (recType.getSelectedItem().equals("===ВЫБЕРИТЕ Тип Записи===")) {
            console.append("В списке RecType не выбрано значение\n");
            return;
        }
        staticData[0] = String.valueOf(recType.getSelectedItem());

        if (lmszid.getSelectedItem().equals("===ВЫБЕРИТЕ ID меры===")) {
            console.append("В списке LMSZID не выбрано значение\n");
            return;
        }
        staticData[1] = String.valueOf(lmszid.getSelectedItem());

        if (categoryID.getSelectedItem().equals("===ВЫБЕРИТЕ ID категории===")) {
            console.append("В списке categoryID не выбрано значение\n");
            return;
        }
        staticData[2] = String.valueOf(categoryID.getSelectedItem());

        if (onmszCode.getSelectedItem().equals("===ВЫБЕРИТЕ Код ОНМСЗ===")) {
            console.append("В списке ONMSZCode не выбрано значение\n");
            return;
        }
        staticData[3] = String.valueOf(onmszCode.getSelectedItem());

        staticData[4] = String.valueOf(usingSign.getSelectedItem());

        staticData[5] = String.valueOf("03");
        staticData[6] = String.valueOf("1");
        staticData[7] = String.valueOf("03");

        staticData[8] = String.valueOf(monetization.getSelectedItem());

        if (decisiondate.getText().trim().isEmpty() || dateStart.getText().trim().isEmpty() || dateFinish.getText().trim().isEmpty()) {
            console.append("Одно из полей с датами не заполнено\n");
            return;
        }
        staticData[9] = String.valueOf(decisiondate.getText().trim());
        staticData[10] = String.valueOf(dateStart.getText().trim());
        staticData[11] = String.valueOf(dateFinish.getText().trim());

        System.out.println(Arrays.toString(staticData));
    }

    private static int countOriginData(File firstFile) throws FileNotFoundException, IOException {
        int countData = 0;
        try (FileInputStream fisFirst = new FileInputStream(firstFile); Workbook firstWorkbook = new XSSFWorkbook(fisFirst)) {

            Sheet firstSheet = firstWorkbook.getSheetAt(0);

            for (Row row : firstSheet) {
                Cell firstCell = row.getCell(0);
                if (firstCell == null || firstCell.getCellType() == CellType.BLANK) {
                    break;
                }
                if (firstCell.getStringCellValue().equals("СНИЛС")) {
                    continue;
                }

                countData++;
            }

        } catch (IOException e) {
        }
        return countData;
    }

    private static int countSecontData(File secondFile) throws FileNotFoundException, IOException {
        int countData = 0;
        try (FileInputStream fisSecond = new FileInputStream(secondFile); Workbook secondWorkbook = new XSSFWorkbook(fisSecond)) {

            Sheet firstSheet = secondWorkbook.getSheetAt(0);

            for (Row row : firstSheet) {
                Cell firstCell = row.getCell(0);
                if (firstCell == null || firstCell.getCellType() == CellType.BLANK) {
                    break;
                }
                countData++;
            }

        } catch (IOException e) {
        }
        return countData;
    }

    public static void mergeData(JTextArea console, File firstFile, File secondFile, String[] staticData) throws FileNotFoundException, IOException {
        try (FileInputStream fisFirst = new FileInputStream(firstFile); FileInputStream fisSecond = new FileInputStream(secondFile); Workbook firstWorkbook = new XSSFWorkbook(fisFirst); Workbook secondWorkbook = new XSSFWorkbook(fisSecond)) {

            Sheet firstSheet = firstWorkbook.getSheetAt(0);
            Sheet secondSheet = secondWorkbook.getSheetAt(0);

            int countOriginRows = countOriginData(firstFile);
            int countSecondRows = countSecontData(secondFile);
            
            System.out.println("В файле с данными строк: "+countOriginRows);
            System.out.println("В шаблоне строк: "+countSecondRows);
            
            int getDataRow = 1;

            for (int i = countSecondRows; i <= countSecondRows + countOriginRows-1; i++) {
                Row targetRow = secondSheet.createRow(i);
                for (int j = getDataRow; j <= getDataRow; j++) {
                    Row sourceRow = firstSheet.getRow(j);

                    Cell snilsSourceSell = sourceRow.getCell(0);
                    Cell lastnameSourceSell = sourceRow.getCell(1);
                    Cell firstnameSourceSell = sourceRow.getCell(2);
                    Cell patronymicSourceSell = sourceRow.getCell(3);
                    Cell genderSourceSell = sourceRow.getCell(4);
                    Cell birthdateSourceSell = sourceRow.getCell(5);
                    Cell amountSourceSell = sourceRow.getCell(6);

                    Cell RecTypeTargetSell = targetRow.createCell(0);
                    Cell LMSZIDTypeTargetSell = targetRow.createCell(2);
                    Cell categoryIDTargetSell = targetRow.createCell(3);
                    Cell ONMSZCodeTargetSell = targetRow.createCell(4);
                    Cell SNILSTargetSell = targetRow.createCell(7);
                    Cell lastnameTargetSell = targetRow.createCell(8);
                    Cell firstnameTargetSell = targetRow.createCell(9);
                    Cell patronymicTargetSell = targetRow.createCell(10);
                    Cell genderTargetSell = targetRow.createCell(11);
                    Cell birthdateTargetSell = targetRow.createCell(12);
                    Cell decisiondateTargetSell = targetRow.createCell(30);
                    Cell dateStartTargetSell = targetRow.createCell(31);
                    Cell dateFinishTargetSell = targetRow.createCell(32);
                    Cell usingSignTargetSell = targetRow.createCell(33);
                    Cell FormCodeTargetSell = targetRow.createCell(36);
                    Cell amountTargetSell = targetRow.createCell(37);
                    Cell measuryCodeTargetSell = targetRow.createCell(38);
                    Cell monetizationTargetSell = targetRow.createCell(39);
                    Cell equivalentAmountTargetSell = targetRow.createCell(42);

                    RecTypeTargetSell.setCellValue(String.valueOf(staticData[0]));
                    LMSZIDTypeTargetSell.setCellValue(String.valueOf(staticData[1]));
                    categoryIDTargetSell.setCellValue(String.valueOf(staticData[2]));
                    ONMSZCodeTargetSell.setCellValue(String.valueOf(staticData[3]));

                    if (snilsSourceSell != null) { //СНИЛС
                        if (snilsSourceSell.getCellType() == CellType.STRING) {
                            SNILSTargetSell.setCellValue(snilsSourceSell.getStringCellValue());
                        } else if (snilsSourceSell.getCellType() == CellType.NUMERIC) {
                            SNILSTargetSell.setCellValue(snilsSourceSell.getNumericCellValue());
                        }
                    }
                    if (lastnameSourceSell != null) {//ФАМИЛИЯ
                        if (lastnameSourceSell.getCellType() == CellType.STRING) {
                            lastnameTargetSell.setCellValue(lastnameSourceSell.getStringCellValue());
                        } else if (lastnameSourceSell.getCellType() == CellType.NUMERIC) {
                            lastnameTargetSell.setCellValue(lastnameSourceSell.getNumericCellValue());
                        }
                    }
                    if (firstnameSourceSell != null) {//ИМЯ
                        if (firstnameSourceSell.getCellType() == CellType.STRING) {
                            firstnameTargetSell.setCellValue(firstnameSourceSell.getStringCellValue());
                        } else if (firstnameSourceSell.getCellType() == CellType.NUMERIC) {
                            firstnameTargetSell.setCellValue(firstnameSourceSell.getNumericCellValue());
                        }
                    }
                    if (patronymicSourceSell != null) {//ОТЧЕСТВО
                        if (patronymicSourceSell.getCellType() == CellType.STRING) {
                            patronymicTargetSell.setCellValue(patronymicSourceSell.getStringCellValue());
                        } else if (patronymicSourceSell.getCellType() == CellType.NUMERIC) {
                            patronymicTargetSell.setCellValue(patronymicSourceSell.getNumericCellValue());
                        }
                    }
                    if (genderSourceSell != null) {//ПОЛ
                        if (genderSourceSell.getCellType() == CellType.STRING) {
                            genderTargetSell.setCellValue(genderSourceSell.getStringCellValue());
                        } else if (genderSourceSell.getCellType() == CellType.NUMERIC) {
                            genderTargetSell.setCellValue(genderSourceSell.getNumericCellValue());
                        }
                    }
                    if (birthdateSourceSell != null) {//ДАТА РОЖДЕНИЯ
                        if (birthdateSourceSell.getCellType() == CellType.STRING) {
                            birthdateTargetSell.setCellValue(birthdateSourceSell.getStringCellValue());
                        } else if (birthdateSourceSell.getCellType() == CellType.NUMERIC) {
                            birthdateTargetSell.setCellValue(birthdateSourceSell.getNumericCellValue());
                        }
                    }

                    decisiondateTargetSell.setCellValue(String.valueOf(staticData[9]));
                    dateStartTargetSell.setCellValue(String.valueOf(staticData[10]));
                    dateFinishTargetSell.setCellValue(String.valueOf(staticData[11]));
                    usingSignTargetSell.setCellValue(String.valueOf(staticData[4]));
                    FormCodeTargetSell.setCellValue(String.valueOf(staticData[5]));
                    amountTargetSell.setCellValue(String.valueOf(staticData[6]));
                    measuryCodeTargetSell.setCellValue(String.valueOf(staticData[7]));
                    monetizationTargetSell.setCellValue(String.valueOf(staticData[8]));

                    if (amountSourceSell != null) {//ДАТА РОЖДЕНИЯ
                        if (amountSourceSell.getCellType() == CellType.STRING) {
                            equivalentAmountTargetSell.setCellValue(amountSourceSell.getStringCellValue());
                        } else if (amountSourceSell.getCellType() == CellType.NUMERIC) {
                            equivalentAmountTargetSell.setCellValue(amountSourceSell.getNumericCellValue());
                        }
                    }
                }
                getDataRow++;
            }

            try (FileOutputStream fos = new FileOutputStream(secondFile)) {
                secondWorkbook.write(fos);
                console.append("Данные успешно объединены и сохранены в файле: " + secondFile.getAbsolutePath() + "\n");
            }
        } catch (IOException e) {
        }
    }

}
