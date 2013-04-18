#excel2rdf

excel2rdf is a java based command-line utility to convent excel file to RDF.  The utility takes two files as input. First an excel file and second an optional  SPARQL query file with a CONSTRUCT query . The result is outputted in TURTLE format.

##Command-line

      excel2rdf table.xls [query.sparql]

  Main arguments

      table.xls              Excel file to be processed
      query.sparql           Optional file containing a SPARQL CONSTRUCT query to be applied to a excel file

  Options

      -v   --verbose         Verbose
      -q   --quiet           Run with minimal output
      --debug                Output information for debugging
      --help
      --version              Version information

the project uses maven as build tool.
