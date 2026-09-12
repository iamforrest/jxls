package org.jxls.transform.poi;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;
import org.junit.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;

public class PoiCellDataDetachedRichTextTest {

    @Test
    public void restoresInlineXssfRichTextAfterTemplateCellIsCleared() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFCell cell = (XSSFCell) workbook.createSheet("template").createRow(0).createCell(0);
            XSSFRichTextString header = new XSSFRichTextString("Static header");
            XSSFFont bold = workbook.createFont();
            bold.setBold(true);
            header.applyFont(0, 6, bold);
            setInlineString(cell, header);

            PoiCellData cellData = createCellData(cell);

            cell.setCellType(CellType.BLANK);
            cellData.writeToCell(cell, new Context(), null);

            XSSFRichTextString restored = cell.getRichStringCellValue();
            assertEquals("Static header", restored.getString());
            assertTrue("rich-text formatting must be preserved", restored.numFormattingRuns() > 0);
            assertEquals(0, restored.getIndexOfFormattingRun(0));
            assertTrue("the first formatting run must stay bold",
                    restored.getFontOfFormattingRun(0).getBold());
        }
    }

    @Test
    public void restoresSharedXssfStringAfterTemplateCellIsCleared() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Cell cell = workbook.createSheet("template").createRow(0).createCell(0);
            XSSFRichTextString header = new XSSFRichTextString("Shared header");
            XSSFFont bold = workbook.createFont();
            bold.setBold(true);
            header.applyFont(0, 6, bold);
            cell.setCellValue(header);
            assertEquals(STCellType.S, ((XSSFCell) cell).getCTCell().getT());

            PoiCellData cellData = createCellData(cell);

            cell.setCellType(CellType.BLANK);
            cellData.writeToCell(cell, new Context(), null);

            XSSFRichTextString restored = (XSSFRichTextString) cell.getRichStringCellValue();
            assertEquals("Shared header", restored.getString());
            assertTrue("rich-text formatting must be preserved", restored.numFormattingRuns() > 0);
            assertTrue("the first formatting run must stay bold",
                    restored.getFontOfFormattingRun(0).getBold());
        }
    }

    @Test
    public void keepsHssfRichTextBehavior() throws Exception {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            Cell cell = workbook.createSheet("template").createRow(0).createCell(0);
            HSSFRichTextString header = new HSSFRichTextString("Legacy header");
            header.applyFont(0, 6, workbook.createFont());
            cell.setCellValue(header);

            PoiCellData cellData = createCellData(cell);

            cell.setCellType(CellType.BLANK);
            cellData.writeToCell(cell, new Context(), null);

            RichTextString restored = cell.getRichStringCellValue();
            assertEquals("Legacy header", restored.getString());
            assertTrue("rich-text formatting must be preserved", restored.numFormattingRuns() > 0);
        }
    }

    private PoiCellData createCellData(Cell cell) {
        PoiRowData rowData = mock(PoiRowData.class);
        when(rowData.getPoiSheetData()).thenReturn(mock(PoiSheetData.class));
        PoiCellData cellData = PoiCellData.createCellData(rowData,
                new CellRef(cell.getSheet().getSheetName(), cell.getRowIndex(), cell.getColumnIndex()), cell);
        Transformer transformer = mock(Transformer.class);
        when(transformer.getTransformationConfig()).thenReturn(new TransformationConfig());
        cellData.setTransformer(transformer);
        return cellData;
    }

    private void setInlineString(XSSFCell cell, XSSFRichTextString value) {
        cell.getCTCell().setT(STCellType.INLINE_STR);
        cell.getCTCell().setIs((CTRst) value.getCTRst().copy());
    }
}
