package com.example.springboot.export;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import java.io.ByteArrayOutputStream;
import java.util.List;

public abstract class ExcelStyle {

    protected Workbook workbook;
    protected ByteArrayOutputStream out;
    protected Sheet sheet;
    protected CellStyle cellStyle;
    protected CellStyle headerCellStyle;
    protected CellStyle cellLeftStyle;
    protected CellStyle hyperLinkCellStype;
    protected CreationHelper creationHelper;
    public ExcelStyle(String sheetName) {
        workbook = new XSSFWorkbook();
        out = new ByteArrayOutputStream();
        sheet = workbook.createSheet(sheetName);
        initStyle();
        sheet.setDefaultRowHeight((short) (420));
        creationHelper= workbook.getCreationHelper();
    }

    private void initStyle() {
        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setFontHeightInPoints((short) 11);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        this.cellStyle = style;

        CellStyle cellLeftStyle = workbook.createCellStyle();
        cellLeftStyle.cloneStyleFrom(style);
//        cellLeftStyle.setWrapText(true);
        cellLeftStyle.setAlignment(HorizontalAlignment.LEFT);
        cellLeftStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        this.cellLeftStyle = cellLeftStyle;

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.cloneStyleFrom(style);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.headerCellStyle = headerCellStyle;

        //Hyperlink cellstyle
        CellStyle hyperLinkCellStype = workbook.createCellStyle();
        hyperLinkCellStype.cloneStyleFrom(style);
        Font hlink_font = workbook.createFont();
        hlink_font.setUnderline(Font.U_SINGLE);
        hlink_font.setColor(IndexedColors.BLUE.getIndex());
        hyperLinkCellStype.setFont(hlink_font);
        this.hyperLinkCellStype = hyperLinkCellStype;
    }

    protected void genHeaderRow(String[] columns) {
        // Header Row
        Row headerRow = sheet.createRow(0);

        // Table Header
        for (int col = 0; col < columns.length; col++) {
            Cell cell = headerRow.createCell(col);
            cell.setCellValue(columns[col]);
            cell.setCellStyle(headerCellStyle);
        }
    }

    protected void genTotalRow(List<String> keySet, JSONObject obj, int dataLength) {
        // Header Row
        Row row = sheet.createRow(dataLength);

        for (int j = 0; j < obj.length(); j++) {
            for (String key : keySet
            ) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(headerCellStyle);
                cell.setCellValue(String.valueOf(obj.get(key)));
                j++;
            }
        }
    }

    protected void autoResizeColumns(int numberOfColumns) {
        for (int i = 0; i < numberOfColumns; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
