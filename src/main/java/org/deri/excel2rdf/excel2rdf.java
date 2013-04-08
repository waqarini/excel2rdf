package org.deri.excel2rdf;

import arq.cmdline.CmdGeneral;
import com.hp.hpl.jena.query.Query;
import com.hp.hpl.jena.query.QueryExecution;
import com.hp.hpl.jena.query.QueryExecutionFactory;
import com.hp.hpl.jena.query.QueryFactory;
import com.hp.hpl.jena.rdf.model.Model;
import com.hp.hpl.jena.rdf.model.ModelFactory;
import com.hp.hpl.jena.shared.NotFoundException;
import com.hp.hpl.jena.sparql.util.Utils;
import org.apache.commons.io.IOUtils;
import java.io.InputStream;
import java.io.FileInputStream;
import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;

public class excel2rdf extends CmdGeneral {

    public static void main(String... args) {
        new excel2rdf(args).mainRun();
    }
    
    private String queryFile;
    private String xlsFile;
    private Model resultModel = ModelFactory.createDefaultModel();

    public excel2rdf(String[] args) {
        super(args);
        getUsage().startCategory("Options");
        getUsage().startCategory("Main arguments");
        getUsage().addUsage("query.sparql", "File containing a SPARQL query to be applied to a excel file");
        getUsage().addUsage("table.xls", "Excel file to be processed");
    }

    @Override
    protected String getCommandName() {
        return Utils.className(this);
    }

    @Override
    protected String getSummary() {
        return getCommandName() + " query.sparql table.xls ";
    }

    @Override
    protected void processModulesAndArgs() {
        if (getPositional().size() < 2) {
            doHelp();
        }
        queryFile = getPositionalArg(0);
        xlsFile = getPositionalArg(1);
    }

    @Override
    protected void exec() {
        try {
            resultModel.setNsPrefix("excel", EXCEL.getURI());
            XLS2RDF xlsToRDF = new XLS2RDF(xlsFile);
            Model model = xlsToRDF.read();
            Query query = QueryFactory.create(loadQuery(queryFile));
            QueryExecution queryExec = QueryExecutionFactory.create(query, model);
            Model results = queryExec.execConstruct();
            //model.write(System.out, "TURTLE");
            results.write(System.out, "TURTLE");
        } catch (Exception ex) {
            Logger.getLogger(excel2rdf.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static String loadQuery(String fileName) throws Exception {
        InputStream inputStream = new FileInputStream(new File(fileName));
        return IOUtils.toString(inputStream, "UTF-8");
    }
}
