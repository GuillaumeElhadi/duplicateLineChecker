package fr.elhadi.ducplicatelinechecker.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

public class ExcelGenerator {

    public static void generateExcelFile(List<Element> elementsANotPresentInB, List<Element> elementsBNotPresentInA) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = prepareSheet(workbook);
        prepareHeader(workbook, sheet);

        CellStyle greenStyle = createStyle(workbook, IndexedColors.LIGHT_GREEN.getIndex());
        CellStyle yellowStyle = createStyle(workbook, IndexedColors.LIGHT_GREEN.getIndex());

        int index = 2;
        for (Element element : elementsANotPresentInB) {
            Optional<Element> elementB = elementsBNotPresentInA.stream().filter(b -> element.getDppType().equals(b.getDppType())
                    && b.getDppThird().equals(element.getDppThird())
                    && b.getDppSubThird().equals(element.getDppSubThird())
                    && b.getCustomerType().equals(element.getCustomerType())
                    && b.getCustomerThird().equals(element.getCustomerThird())
                    && b.getCustomerSubTihrd().equals(element.getCustomerSubTihrd())
                    && b.getFinalType().equals(element.getFinalType())
                    && b.getFinalThird().equals(element.getFinalThird())
                    && b.getFinalSubThird().equals(element.getFinalSubThird())
                    && b.getItemNumber().equals(element.getItemNumber())).findFirst();
            Row row = sheet.createRow(index);
            fillCells(sheet, greenStyle, yellowStyle, index, element, elementB, row);

            index++;
        }
        saveExcelFile(workbook);
    }

    private static CellStyle createStyle(Workbook workbook, short color) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("@"));
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        return style;
    }

    private static void fillCells(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent()) {
            fillDppType(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillDppThird(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillDppSubThird(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillCustomerType(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillCustomerThird(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillCustomerSubTird(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillFinalType(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillFinalThird(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillFinalSubThird(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillItemNumber(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillCurrentWeek(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillCurrentYear(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillReplenishment(sheet, greenStyle, yellowStyle, index, element, elementB, row);
            fillProvisions(sheet, greenStyle, yellowStyle, index, element, elementB, row);
        }
    }

    private static void fillProvisions(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getProvisions().equals(elementB.get().getProvisions())) {
            Cell originalProvisions = row.createCell(26);
            originalProvisions.setCellValue(element.getProvisions());
            originalProvisions.setCellStyle(yellowStyle);

            Cell optimizedProvisions = row.createCell(27);
            optimizedProvisions.setCellValue(elementB.get().getProvisions());
            optimizedProvisions.setCellStyle(yellowStyle);

        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 26, 27));
            Cell originalProvisions = row.createCell(26);
            originalProvisions.setCellValue(element.getProvisions());
            originalProvisions.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(27);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillReplenishment(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getReplenishment().equals(elementB.get().getReplenishment())) {
            Cell originalReplenishment = row.createCell(24);
            originalReplenishment.setCellValue(element.getReplenishment());
            originalReplenishment.setCellStyle(yellowStyle);

            Cell optimizedReplenishment = row.createCell(25);
            optimizedReplenishment.setCellValue(elementB.get().getReplenishment());
            optimizedReplenishment.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 24, 25));
            Cell originalReplenishment = row.createCell(24);
            originalReplenishment.setCellValue(element.getReplenishment());
            originalReplenishment.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(25);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillCurrentYear(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getCurrentYear().equals(elementB.get().getCurrentYear())) {
            Cell originalCurrentYear = row.createCell(22);
            originalCurrentYear.setCellValue(element.getCurrentYear());
            originalCurrentYear.setCellStyle(yellowStyle);

            Cell optimizedCurrentYear = row.createCell(23);
            optimizedCurrentYear.setCellValue(elementB.get().getCurrentYear());
            optimizedCurrentYear.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 22, 23));
            Cell originalCurrentYear = row.createCell(22);
            originalCurrentYear.setCellValue(element.getCurrentYear());
            originalCurrentYear.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(23);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillCurrentWeek(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getCurrentWeek().equals(elementB.get().getCurrentWeek())) {
            Cell originalCurrentWeek = row.createCell(20);
            originalCurrentWeek.setCellValue(element.getCurrentWeek());
            originalCurrentWeek.setCellStyle(yellowStyle);

            Cell optimizedCurrentWeek = row.createCell(21);
            optimizedCurrentWeek.setCellValue(elementB.get().getCurrentWeek());
            optimizedCurrentWeek.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 20, 21));
            Cell originalCurrentWeek = row.createCell(20);
            originalCurrentWeek.setCellValue(element.getCurrentWeek());
            originalCurrentWeek.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(21);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillItemNumber(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getItemNumber().equals(elementB.get().getItemNumber())) {
            Cell originalItemNumber = row.createCell(18);
            originalItemNumber.setCellValue(element.getItemNumber());
            originalItemNumber.setCellStyle(yellowStyle);

            Cell optimizedItemNumber = row.createCell(19);
            optimizedItemNumber.setCellValue(elementB.get().getItemNumber());
            optimizedItemNumber.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 18, 19));
            Cell originalItemNumber = row.createCell(18);
            originalItemNumber.setCellValue(element.getItemNumber());
            originalItemNumber.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(19);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillFinalSubThird(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getFinalSubThird().equals(elementB.get().getFinalSubThird())) {
            Cell originalFinalSubThird = row.createCell(16);
            originalFinalSubThird.setCellValue(element.getFinalSubThird());
            originalFinalSubThird.setCellStyle(yellowStyle);

            Cell optimizedFinalSubThird = row.createCell(17);
            optimizedFinalSubThird.setCellValue(elementB.get().getFinalSubThird());
            optimizedFinalSubThird.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 16, 17));
            Cell originalFinalSubThird = row.createCell(16);
            originalFinalSubThird.setCellValue(element.getFinalSubThird());
            originalFinalSubThird.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(17);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillFinalThird(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getFinalThird().equals(elementB.get().getFinalThird())) {
            Cell originalFinalThird = row.createCell(14);
            originalFinalThird.setCellValue(element.getFinalThird());
            originalFinalThird.setCellStyle(yellowStyle);

            Cell optimizedFinalThird = row.createCell(15);
            optimizedFinalThird.setCellValue(elementB.get().getFinalThird());
            optimizedFinalThird.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 14, 15));
            Cell originalFinalThird = row.createCell(14);
            originalFinalThird.setCellValue(element.getFinalThird());
            originalFinalThird.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(15);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillFinalType(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() &&  !element.getFinalType().equals(elementB.get().getFinalType())) {
            Cell originalFinalType = row.createCell(12);
            originalFinalType.setCellValue(element.getFinalType());
            originalFinalType.setCellStyle(yellowStyle);

            Cell optimizedFinalType = row.createCell(13);
            optimizedFinalType.setCellValue(elementB.get().getFinalType());
            optimizedFinalType.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 12, 13));
            Cell originalFinalType = row.createCell(12);
            originalFinalType.setCellValue(element.getFinalType());
            originalFinalType.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(13);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillCustomerSubTird(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getCustomerSubTihrd().equals(elementB.get().getCustomerSubTihrd())) {
            Cell originalCustomerSubThird = row.createCell(10);
            originalCustomerSubThird.setCellValue(element.getCustomerSubTihrd());
            originalCustomerSubThird.setCellStyle(yellowStyle);

            Cell optimizedCustomerSubThird = row.createCell(11);
            optimizedCustomerSubThird.setCellValue(elementB.get().getCustomerSubTihrd());
            optimizedCustomerSubThird.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 10, 11));
            Cell originalCustomerSubThird = row.createCell(10);
            originalCustomerSubThird.setCellValue(element.getCustomerSubTihrd());
            originalCustomerSubThird.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(11);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillCustomerThird(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getCustomerThird().equals(elementB.get().getCustomerThird())) {
            Cell originalCustomerThird = row.createCell(8);
            originalCustomerThird.setCellValue(element.getCustomerThird());
            originalCustomerThird.setCellStyle(yellowStyle);

            Cell optimizedCustomerThird = row.createCell(9);
            optimizedCustomerThird.setCellValue(elementB.get().getCustomerThird());
            optimizedCustomerThird.setCellStyle(yellowStyle);

        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 8, 9));
            Cell originalCustomerThird = row.createCell(8);
            originalCustomerThird.setCellValue(element.getCustomerThird());
            originalCustomerThird.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(9);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillCustomerType(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getCustomerType().equals(elementB.get().getCustomerType())) {
            Cell originalCustomerType = row.createCell(6);
            originalCustomerType.setCellValue(element.getCustomerType());
            originalCustomerType.setCellStyle(yellowStyle);

            Cell optimizedCustomerType = row.createCell(7);
            optimizedCustomerType.setCellValue(elementB.get().getCustomerType());
            optimizedCustomerType.setCellStyle(yellowStyle);

        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 6, 7));
            Cell originalCustomerType = row.createCell(6);
            originalCustomerType.setCellValue(element.getCustomerType());
            originalCustomerType.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(7);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillDppSubThird(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getDppSubThird().equals(elementB.get().getDppSubThird())) {
            Cell originalDppSubThird = row.createCell(4);
            originalDppSubThird.setCellValue(element.getDppSubThird());
            originalDppSubThird.setCellStyle(yellowStyle);

            Cell optimizedDppSubThird = row.createCell(5);
            optimizedDppSubThird.setCellValue(elementB.get().getDppSubThird());
            optimizedDppSubThird.setCellStyle(yellowStyle);

        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 4, 5));
            Cell originalDppSubThird = row.createCell(4);
            originalDppSubThird.setCellValue(element.getDppSubThird());
            originalDppSubThird.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(5);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillDppThird(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getDppThird().equals(elementB.get().getDppThird())) {
            Cell originalDppThird = row.createCell(2);
            originalDppThird.setCellValue(element.getDppThird());
            originalDppThird.setCellStyle(yellowStyle);

            Cell optimizedDppThird = row.createCell(3);
            optimizedDppThird.setCellValue(elementB.get().getDppThird());
            optimizedDppThird.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 2, 3));
            Cell originalDppThird = row.createCell(2);
            originalDppThird.setCellValue(element.getDppThird());
            originalDppThird.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(3);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void fillDppType(Sheet sheet, CellStyle greenStyle, CellStyle yellowStyle, int index, Element element, Optional<Element> elementB, Row row) {
        if (elementB.isPresent() && !element.getDppType().equals(elementB.get().getDppType())) {
            Cell originalDppType = row.createCell(0);
            originalDppType.setCellValue(element.getDppType());
            originalDppType.setCellStyle(yellowStyle);

            Cell optimizedlDppType = row.createCell(1);
            optimizedlDppType.setCellValue(elementB.get().getDppType());
            optimizedlDppType.setCellStyle(yellowStyle);
        } else {
            sheet.addMergedRegion(new CellRangeAddress(index, index, 0, 1));
            Cell originalDppType = row.createCell(0);
            originalDppType.setCellValue(element.getDppType());
            originalDppType.setCellStyle(greenStyle);

            Cell emptyCell = row.createCell(1);
            emptyCell.setCellStyle(greenStyle);
        }
    }

    private static void saveExcelFile(Workbook workbook) throws IOException {
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();
    }

    private static Sheet prepareSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("Items difference");
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 4000);
        sheet.setColumnWidth(2, 4000);
        sheet.setColumnWidth(3, 4000);
        sheet.setColumnWidth(4, 4000);
        sheet.setColumnWidth(5, 4000);
        sheet.setColumnWidth(6, 4000);
        sheet.setColumnWidth(7, 4000);
        sheet.setColumnWidth(8, 4000);
        sheet.setColumnWidth(9, 4000);
        sheet.setColumnWidth(10, 4000);
        sheet.setColumnWidth(11, 4000);
        sheet.setColumnWidth(12, 4000);
        sheet.setColumnWidth(13, 4000);
        sheet.setColumnWidth(14, 4000);
        sheet.setColumnWidth(15, 4000);
        sheet.setColumnWidth(16, 4000);
        sheet.setColumnWidth(17, 4000);
        sheet.setColumnWidth(18, 4000);
        sheet.setColumnWidth(19, 4000);
        sheet.setColumnWidth(20, 4000);
        sheet.setColumnWidth(21, 4000);
        sheet.setColumnWidth(22, 4000);
        sheet.setColumnWidth(23, 4000);
        sheet.setColumnWidth(24, 4000);
        sheet.setColumnWidth(25, 4000);
        sheet.setColumnWidth(26, 4000);
        sheet.setColumnWidth(27, 4000);
        return sheet;
    }

    private static void prepareHeader(Workbook workbook, Sheet sheet) {
        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 12);
        headerStyle.setFont(font);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("DppType");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 3));
        headerCell = header.createCell(2);
        headerCell.setCellValue("DppThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 5));
        headerCell = header.createCell(4);
        headerCell.setCellValue("DppSubThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 7));
        headerCell = header.createCell(6);
        headerCell.setCellValue("CustomerType");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 8, 9));
        headerCell = header.createCell(8);
        headerCell.setCellValue("CustomerThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 10, 11));
        headerCell = header.createCell(10);
        headerCell.setCellValue("CustomerSubThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 12, 13));
        headerCell = header.createCell(12);
        headerCell.setCellValue("FinalType");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 14, 15));
        headerCell = header.createCell(14);
        headerCell.setCellValue("FinalThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 16, 17));
        headerCell = header.createCell(16);
        headerCell.setCellValue("FinalSubThird");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 18, 19));
        headerCell = header.createCell(18);
        headerCell.setCellValue("ItemNumber");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 20, 21));
        headerCell = header.createCell(20);
        headerCell.setCellValue("CurrentWeek");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 22, 23));
        headerCell = header.createCell(22);
        headerCell.setCellValue("CurrentYear");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 24, 25));
        headerCell = header.createCell(24);
        headerCell.setCellValue("replenishment");
        headerCell.setCellStyle(headerStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 26, 27));
        headerCell = header.createCell(26);
        headerCell.setCellValue("provisions");
        headerCell.setCellStyle(headerStyle);

        Row subHeader = sheet.createRow(1);
        for (int i = 0; i < 28; i += 2) {
            headerCell = subHeader.createCell(i);
            headerCell.setCellValue("original");
            headerCell.setCellStyle(headerStyle);

            headerCell = subHeader.createCell(i + 1);
            headerCell.setCellValue("optimized");
            headerCell.setCellStyle(headerStyle);

        }
    }
}
