package org.jxls.transform.poi;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * POI utility methods
 *
 * @author Leonid Vysochyn
 */
public class PoiUtil {

    public static void setCellComment(Cell cell, String commentText, String commentAuthor, ClientAnchor anchor) {
        Sheet sheet = cell.getSheet();
        Workbook wb = sheet.getWorkbook();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        if (anchor == null) {
            anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex() + 1);
            anchor.setCol2(cell.getColumnIndex() + 3);
            anchor.setRow1(cell.getRowIndex());
            anchor.setRow2(cell.getRowIndex() + 2);
        }
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(commentAuthor != null ? commentAuthor : "");
        cell.setCellComment(comment);
    }

    public WritableCellValue hyperlink(String address, String link, String linkTypeString) {
        return new WritableHyperlink(address, link, linkTypeString);
    }

    public WritableCellValue hyperlink(String address, String title) {
        return new WritableHyperlink(address, title);
    }

    public static void copySheetProperties(Sheet src, Sheet dest) {
        dest.setAutobreaks(src.getAutobreaks());
        dest.setDisplayGridlines(src.isDisplayGridlines());
        dest.setVerticallyCenter(src.getVerticallyCenter());
        dest.setFitToPage(src.getFitToPage());
        dest.setForceFormulaRecalculation(src.getForceFormulaRecalculation());
        dest.setRowSumsRight(src.getRowSumsRight());
        dest.setRowSumsBelow(src.getRowSumsBelow());
        copyPrintSetup(src, dest);
    }

    private static void copyPrintSetup(Sheet src, Sheet dest) {
        PrintSetup srcPrintSetup = src.getPrintSetup();
        PrintSetup destPrintSetup = dest.getPrintSetup();
        destPrintSetup.setCopies(srcPrintSetup.getCopies());
        destPrintSetup.setDraft(srcPrintSetup.getDraft());
        destPrintSetup.setFitHeight(srcPrintSetup.getFitHeight());
        destPrintSetup.setFitWidth(srcPrintSetup.getFitWidth());
        destPrintSetup.setFooterMargin(srcPrintSetup.getFooterMargin());
        destPrintSetup.setHeaderMargin(srcPrintSetup.getHeaderMargin());
        destPrintSetup.setHResolution(srcPrintSetup.getHResolution());
        destPrintSetup.setLandscape(srcPrintSetup.getLandscape());
        destPrintSetup.setLeftToRight(srcPrintSetup.getLeftToRight());
        destPrintSetup.setNoColor(srcPrintSetup.getNoColor());
        destPrintSetup.setNoOrientation(srcPrintSetup.getNoOrientation());
        destPrintSetup.setNotes(srcPrintSetup.getNotes());
        destPrintSetup.setPageStart(srcPrintSetup.getPageStart());
        destPrintSetup.setPaperSize(srcPrintSetup.getPaperSize());
        destPrintSetup.setScale(srcPrintSetup.getScale());
        destPrintSetup.setUsePage(srcPrintSetup.getUsePage());
        destPrintSetup.setValidSettings(srcPrintSetup.getValidSettings());
        destPrintSetup.setVResolution(srcPrintSetup.getVResolution());
    }

    public static boolean isJxComment(String cellComment) {
        if (cellComment == null) return false;
        String[] commentLines = cellComment.split("\\n");
        for (String commentLine : commentLines) {
            if ((commentLine != null) && XlsCommentAreaBuilder.isCommandString(commentLine.trim())) {
                return true;
            }
        }
        return false;
    }

    public static void copyImages(XSSFSheet from, XSSFSheet to) throws Exception {
        Drawing drawingPatriarch = to.createDrawingPatriarch();
        XSSFWorkbook destWorkbook = to.getWorkbook();

        // Add image
        for (POIXMLDocumentPart pdp : from.getRelations()) {
            if (!(pdp instanceof XSSFDrawing)) {
                continue;
            }
            List<XSSFShape> shapes = ((XSSFDrawing) pdp).getShapes();
            AtomicInteger atomicInteger = new AtomicInteger();
            for (XSSFShape shape : shapes) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture srcPic = (XSSFPicture) shape;
                    ClientAnchor srcAnchor = srcPic.getClientAnchor();

                    ClientAnchor destAnchor = destWorkbook.getCreationHelper().createClientAnchor();
                    BeanUtils.copyProperties(destAnchor, srcAnchor);
                    Picture picture = drawingPatriarch.createPicture(srcAnchor, atomicInteger.getAndIncrement());
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        Context jxlsContext = new Context();
        jxlsContext.putVar("rows", new ArrayList<>());

        FileInputStream fis = new FileInputStream("/Users/forrest/Downloads/template.xlsx");
        FileOutputStream fos = new FileOutputStream("/Users/forrest/Downloads/template.out.xlsx");

        // image lost if true
        boolean useStreaming = true;
        if (useStreaming) {
            Workbook workbook = WorkbookFactory.create(fis);
//            SXSSFWorkbook xssfWorkbook = new SXSSFWorkbook((XSSFWorkbook) workbook, 600, false, false);
//            xssfWorkbook.write(fos);

            PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 600, false);

            AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
            List<Area> xlsAreaList = areaBuilder.build();
            Area xlsArea = xlsAreaList.get(0);
            xlsArea.applyAt(new CellRef("Result!A1"), jxlsContext);
            xlsArea.processFormulas();
            jxlsContext.getConfig().setIsFormulaProcessingRequired(false); // with SXSSF you cannot use normal formula
            jxlsContext.getConfig().setIgnoreSourceCellStyle(true);
            workbook.setForceFormulaRecalculation(true);
            workbook.setActiveSheet(1);
            List<String> templateSheetsName = new ArrayList<>();
            for (Area xlsArea2 : xlsAreaList) {
                templateSheetsName.add(xlsArea2.getAreaRef().getSheetName());
            }

            for (String sheetName : templateSheetsName) {
                transformer.deleteSheet(sheetName);
            }
            transformer.setIgnoreColumnProps(true);
            transformer.setIgnoreRowProps(true);
            transformer.getWorkbook().setSheetOrder("Result", 0);
            transformer.getWorkbook().setActiveSheet(0);
            transformer.setOutputStream(fos);
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            jxlsHelper.processTemplate(jxlsContext, transformer);
        } else {
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            Transformer transformer = jxlsHelper.createTransformer(fis, fos);
            jxlsHelper.processTemplate(jxlsContext, transformer);
        }
    }

}
