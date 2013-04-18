#excel2rdf

excel2rdf is a java based command-line utility to convent excel file to RDF.  The utility takes two files as input. First a SPARQL query file with a CONSTRUCT query and second an excel file. The result is outputted in TURTLE format.

##Command-line

    excel2rdf query.sparql table.xls 

  Main arguments

      query.sparql           File containing a SPARQL query to be applied to a excel file
      table.xls              Excel file to be processed

  Options

      -v   --verbose         Verbose

      -q   --quiet           Run with minimal output

      --debug                Output information for debugging

      --help

      --version              Version information


the project uses maven as build tool.
