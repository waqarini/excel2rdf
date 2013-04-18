package org.deri.excel2rdf;

import org.deri.excel2rdf.vocab.EXCEL;
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
   
    public excel2rdf(String[] args) {
        super(args);
        getUsage().startCategory("Options");
        getUsage().startCategory("Main arguments");
        getUsage().addUsage("table.xls", "Excel file to be processed");
        getUsage().addUsage("query.sparql", "Optional file containing a SPARQL CONSTRUCT query to be applied to a excel file");

    }

    @Override
    protected String getCommandName() {
        return Utils.className(this);
    }

    @Override
    protected String getSummary() {
        return getCommandName() + " table.xls [query.sparql]";
    }

    @Override
    protected void processModulesAndArgs() {
        if (getPositional().size() < 1) {
            doHelp();
        }
        if (getPositional().size() == 2) {
            queryFile = getPositionalArg(1);
            xlsFile = getPositionalArg(0);
        } else {
            queryFile = null;
            xlsFile = getPositionalArg(0);
        }
    }

    @Override
    protected void exec() {
        try {
            XLS2RDF xlsToRDF = new XLS2RDF(xlsFile);
            Model model = xlsToRDF.read();
            if (queryFile != null) {
                Query query = QueryFactory.create(loadQuery(queryFile));
                QueryExecution queryExec = QueryExecutionFactory.create(query, model);
                Model results = queryExec.execConstruct();
                results.write(System.out, "TURTLE");
            }else {
                model.write(System.out, "TURTLE");
            }
            
        } catch (Exception ex) {
            Logger.getLogger(excel2rdf.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static String loadQuery(String fileName) throws Exception {
        InputStream inputStream = new FileInputStream(new File(fileName));
        return IOUtils.toString(inputStream, "UTF-8");
    }
}
