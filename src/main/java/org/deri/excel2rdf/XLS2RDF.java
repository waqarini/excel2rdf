package org.deri.excel2rdf;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import com.hp.hpl.jena.graph.Node;
import com.hp.hpl.jena.rdf.model.Model;
import com.hp.hpl.jena.rdf.model.ModelFactory;
import com.hp.hpl.jena.rdf.model.Resource;

import com.hp.hpl.jena.vocabulary.RDF;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class XLS2RDF {

    private final String fileName;

    public XLS2RDF(String fileName) {
        this.fileName = fileName;
    }

    public Model read() {
        Model model = ModelFactory.createDefaultModel();
        try {
            FileInputStream xlFile = new FileInputStream(new File(fileName));
            HSSFWorkbook workbook = new HSSFWorkbook(xlFile);
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                processSheet(workbook.getSheetAt(sheetIndex),model);
            }
            xlFile.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return model;
    }

    private static void processCell(Cell cell, Model model) {
        CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        String sheetName = cell.getSheet().getSheetName();
        String value = cell.toString();
        String up = up(cell);
        String down = down(cell);
        String left = left(cell);
        String right = right(cell);
        String ref = cellRef.formatAsString().replace("$", "");
        String bgColor = ((HSSFColor) cell.getCellStyle().getFillBackgroundColorColor()).getHexString();
        String fgColor = ((HSSFColor) cell.getCellStyle().getFillForegroundColorColor()).getHexString();
        Resource resource = model.createResource(sheetName + "#" + ref);
        resource.addProperty(RDF.type, EXCEL.cell);
        // resource.addProperty(EXCEL.row, );
        //resource.addProperty(EXCEL.column, );
        resource.addProperty(EXCEL.value, value);
        resource.addProperty(EXCEL.sheet, model.createResource(sheetName));
        resource.addProperty(EXCEL.value, value);
        resource.addProperty(EXCEL.foregroundColor, fgColor);
        resource.addProperty(EXCEL.backgroundColor, bgColor);
        if(up!=null)resource.addProperty(EXCEL.up, model.createResource(sheetName+"#"+up));
        if(down!=null)resource.addProperty(EXCEL.down,model.createResource(sheetName+"#"+ down));
        if(left!=null)resource.addProperty(EXCEL.left,model.createResource(sheetName+"#"+left));
        if(right!=null)resource.addProperty(EXCEL.right,model.createResource(sheetName+"#"+right));

    }

    private static String down(Cell cell) {
        try {
            CellReference cellRef = new CellReference(cell.getRowIndex() + 1, cell.getColumnIndex());
            return cellRef.formatAsString().replace("$", "");
        } catch (Exception e) {
        }
        return null;
    }

    private static String up(Cell cell) {
        try {
            CellReference cellRef = new CellReference(cell.getRowIndex() - 1, cell.getColumnIndex());
            if (cellRef.getCol() > 0) {
                return cellRef.formatAsString().replace("$", "");
            }
        } catch (Exception e) {
        }
        return null;
    }

    private static String left(Cell cell) {
        try {
            CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex() - 1);
            if (cellRef.getRow() > 0) {
                return cellRef.formatAsString().replace("$", "");
            }
        } catch (Exception e) {
        }
        return null;
    }

    private static String right(Cell cell) {
        try {
            CellReference cellRef = new CellReference(cell.getRowIndex(), cell.getColumnIndex() + 1);
            return cellRef.formatAsString().replace("$", "");
        } catch (Exception e) {
        }
        return null;
    }

    private static void processSheet(HSSFSheet sheet, Model model) {
        Iterator<Row> rowIterator = sheet.iterator();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {

                processCell(cellIterator.next(), model);

            }

        }
    }
}
