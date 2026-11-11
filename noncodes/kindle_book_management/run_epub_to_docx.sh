#!/usr/bin/env bash
# Runs the Groovy epub -> docx converter without Maven. Jars come from ~/.m2 (versions as in kndiy-tools pom).
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
R=~/.m2/repository
CP="$R/org/apache/groovy/groovy/4.0.23/groovy-4.0.23.jar:$R/org/apache/groovy/groovy-xml/4.0.23/groovy-xml-4.0.23.jar"
for j in org/apache/poi/poi/5.2.5/poi-5.2.5.jar org/apache/poi/poi-ooxml/5.2.5/poi-ooxml-5.2.5.jar \
  org/apache/poi/poi-ooxml-lite/5.2.5/poi-ooxml-lite-5.2.5.jar org/apache/xmlbeans/xmlbeans/5.2.0/xmlbeans-5.2.0.jar \
  org/apache/commons/commons-collections4/4.4/commons-collections4-4.4.jar org/apache/commons/commons-compress/1.26.2/commons-compress-1.26.2.jar \
  commons-io/commons-io/2.15.1/commons-io-2.15.1.jar org/apache/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar \
  com/zaxxer/SparseBitSet/1.3/SparseBitSet-1.3.jar org/apache/logging/log4j/log4j-api/2.21.1/log4j-api-2.21.1.jar; do
  CP="$CP:$R/$j"
done
exec java -cp "$CP:$HERE/src" groovy.ui.GroovyMain "$HERE/run_epub_to_docx.groovy" "$@"
