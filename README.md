#excel2rdf

excel2rdf is a java based command-line utility to convent excel file to RDF.  The utility takes two files as input. First an excel file and second an optional  SPARQL query file with a CONSTRUCT query . The result is outputted in TURTLE format.

##Command-line

  excel2rdf --input table.xls

  Options
      --input                excel file to be used input
      --enrich               Optional SPARQL file to be used for enrichment
      --query                Optional SPARQL file with SELECT query
      --construct            Optional SPARQL file with CONSTRUCT query
  Symbol definition
      --set                  Set a configuration symbol to a value
      --strict               Operate in strict SPARQL mode (no extensions of any kind)
  General
      -v   --verbose         Verbose
      -q   --quiet           Run with minimal output
      --debug                Output information for debugging
      --help
      --version              Version information

the project uses maven as build tool.

To create executable scripts for Windows and Unix in `/target/appassembler/bin/excel2rdf`:

    mvn package appassembler:assemble

