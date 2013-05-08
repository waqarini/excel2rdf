package org.deri.excel2rdf;

import arq.cmdline.ArgDecl;
import arq.cmdline.CmdARQ;
import com.hp.hpl.jena.query.Query;
import com.hp.hpl.jena.query.QueryExecution;
import com.hp.hpl.jena.query.QueryExecutionFactory;
import com.hp.hpl.jena.query.QueryFactory;
import com.hp.hpl.jena.query.ResultSet;
import com.hp.hpl.jena.query.ResultSetFormatter;
import com.hp.hpl.jena.rdf.model.Model;
import com.hp.hpl.jena.rdf.model.ResIterator;
import com.hp.hpl.jena.rdf.model.Resource;
import com.hp.hpl.jena.rdf.model.Statement;
import com.hp.hpl.jena.sparql.util.Utils;
import com.hp.hpl.jena.vocabulary.RDF;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.util.CellReference;
import org.deri.excel2rdf.vocab.EXCEL;

public class excel2rdf extends CmdARQ {

    protected ArgDecl enrichDecl = new ArgDecl(ArgDecl.HasValue, "enrich");
    protected ArgDecl excelDecl = new ArgDecl(ArgDecl.HasValue, "input");
    protected ArgDecl queryDecl = new ArgDecl(ArgDecl.HasValue, "query");
    protected ArgDecl constructDecl = new ArgDecl(ArgDecl.HasValue, "construct");

    public static void main(String... args) {
        new excel2rdf(args).mainRun();
    }
    private String queryFile;
    private String inputFile;
    private String enrichFile;
    private String constructFile;
    private final int MAX_ROWS = 65536;
    private final int MAX_COLS = 256;

    public excel2rdf(String[] args) {
        super(args);
        getUsage().startCategory("Options");
        add(excelDecl, "--input", "excel file to be used input");
        add(enrichDecl, "--enrich", "Optional SPARQL file to be used for enrichment");
        add(queryDecl, "--query", "Optional SPARQL file with SELECT query");
        add(constructDecl, "--construct", "Optional SPARQL file with CONSTRUCT query");

    }

    @Override
    protected String getCommandName() {
        return Utils.className(this);
    }

    @Override
    protected String getSummary() {
        return getCommandName() + " --input table.xls";
    }

    @Override
    protected void processModulesAndArgs() {
        super.processModulesAndArgs();
        if (contains(excelDecl)) {
            inputFile = getValue(excelDecl);
        } else {
            doHelp();
            return;
        }
        if (contains(enrichDecl)) {
            enrichFile = getValue(enrichDecl);
        }

        if (contains(queryDecl)) {
            queryFile = getValue(queryDecl);
        }

        if (contains(constructDecl)) {
            constructFile = getValue(constructDecl);
        }


    }

    @Override
    protected void exec() {
        try {

            XLS2RDF xlsToRDF = new XLS2RDF(inputFile);
            Model model = xlsToRDF.read();


            if (enrichFile != null) {
                enrich(model, enrichFile);
            }
            
            if (constructFile != null) {
                Query query = QueryFactory.create(loadQuery(constructFile));
                QueryExecution queryExec = QueryExecutionFactory.create(query, model);
                Model results = queryExec.execConstruct();
                results.write(System.out, "TURTLE");
                return ;
            }

            if (queryFile != null) {
                Query query = QueryFactory.create(loadQuery(queryFile));
                QueryExecution queryExec = QueryExecutionFactory.create(query, model);
                ResultSet result = queryExec.execSelect();
                ResultSetFormatter.out(System.out, result);
            } else {
                model.write(System.out, "TURTLE");
            }

        } catch (Exception ex) {
            Logger.getLogger(excel2rdf.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private String getCellRef(int row, int col) {
        CellReference cellRef = new CellReference(row - 1, col);
        return cellRef.formatAsString().replace("$", "");

    }

    private void enrich(Model model, String file) throws Exception {

        Query query = QueryFactory.create(loadQuery(file));
        QueryExecution queryExec = QueryExecutionFactory.create(query, model);
        Model enrichModel = queryExec.execConstruct();
        model.add(enrichModel);

        ResIterator sheets = model.listResourcesWithProperty(RDF.type, EXCEL.Sheet);
        while (sheets.hasNext()) {
            Resource sheet = sheets.nextResource();
            //System.out.println(sheet);
            ResIterator cells = model.listResourcesWithProperty(EXCEL.sheet, sheet);
            while (cells.hasNext()) {
                Resource cell = cells.nextResource();
                if (cell.getProperty(EXCEL.cellType) != null) {

                    String col = cell.getProperty(EXCEL.column).getString();
                    int row = cell.getProperty(EXCEL.row).getInt();
                    int icol = cell.getProperty(EXCEL.columnIndex).getInt();
                    String type = cell.getProperty(EXCEL.cellType).getString();

                    cell.addProperty(EXCEL.findNextUp, cell);
                    cell.addProperty(EXCEL.findNextDown, cell);
                    cell.addProperty(EXCEL.findNextLeft, cell);
                    cell.addProperty(EXCEL.findNextRight, cell);

                    //findNextUp
                    int r = row + 1;
                    while (r < MAX_ROWS) {
                        Resource resource = model.getResource(sheet.getLocalName() + "#" + getCellRef(r, icol));
                        Statement t = resource.getProperty(EXCEL.cellType);
                        if (t != null && !resource.getProperty(EXCEL.cellType).equals(type)) {
                            break;
                        }
                        resource.addProperty(EXCEL.findNextUp, cell);
                        r++;
                    }

                    //findNextDown
                    r = row - 1;
                    while (r > 0) {
                        Resource resource = model.getResource(sheet.getLocalName() + "#" + getCellRef(r, icol));
                        Statement t = resource.getProperty(EXCEL.cellType);
                        if (t != null && !resource.getProperty(EXCEL.cellType).equals(type)) {
                            break;
                        }
                        resource.addProperty(EXCEL.findNextDown, cell);
                        r--;
                    }

                    //findNextLeft
                    int c = icol - 1;
                    while (c > 0) {
                        Resource resource = model.getResource(sheet.getLocalName() + "#" + getCellRef(r, icol));
                        Statement t = resource.getProperty(EXCEL.cellType);
                        if (t != null && !resource.getProperty(EXCEL.cellType).equals(type)) {
                            break;
                        }
                        resource.addProperty(EXCEL.findNextLeft, cell);
                        c--;
                    }

                    //findNextDown
                    c = icol + 1;
                    while (c < MAX_COLS) {
                        Resource resource = model.getResource(sheet.getLocalName() + "#" + getCellRef(r, icol));
                        Statement t = resource.getProperty(EXCEL.cellType);
                        if (t != null && !resource.getProperty(EXCEL.cellType).equals(type)) {
                            break;
                        }
                        resource.addProperty(EXCEL.findNextRight, cell);
                        c++;
                    }


                }
            }
        }
    }

    private static String loadQuery(String fileName) throws Exception {
        InputStream inputStream = new FileInputStream(new File(fileName));
        return IOUtils.toString(inputStream, "UTF-8");
    }
}
