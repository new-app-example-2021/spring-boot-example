package com.example.springboot.export;

import com.example.springboot.entity.Card;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.ByteArrayInputStream;
import java.util.List;

public class ExportExcel extends ExcelStyle {

    public ExportExcel() {
        super("export");
    }

    public ByteArrayInputStream exportToExcel(List<Card> cards) {
        String[] COLUMNs = {"Card Name", "User Count", "App Card Send Count", "API Count"};

        try {
            // Header Row
            genHeaderRow(COLUMNs);

            int rowIdx = 1;
            for (Card card : cards) {
                Row row = sheet.createRow(rowIdx++);

                String[] results = {card.getName(), String.valueOf(card.getUserCount()), String.valueOf(card.getSentTimeCount()), String.valueOf(card.getApiCount())};

                for (int i = 0; i < results.length; i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(results[i]);
                }
            }

            //Auto-size all the above columns
            autoResizeColumns(COLUMNs.length);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
