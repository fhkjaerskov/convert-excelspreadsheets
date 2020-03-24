#!/bin/bash 

#Rydder op
echo "Cleaning up"

rm -r /mnt/c/Regneark_fra_s_drevet/konverteringstest/_XLS/
rm -r /mnt/c/Regneark_fra_s_drevet/konverteringstest/_XLSX/
rm -r /mnt/c/Regneark_fra_s_drevet/konverteringstest/_ODS/
rm /mnt/c/Regneark_fra_s_drevet/konverteringstest/logorig*.xml
mkdir _XLS
mkdir _XLSX
mkdir _ODS

#Konverter til ODS
echo "Konverterer Excelfiler til ODS"
cd /mnt/c/Regneark_fra_s_drevet/konverteringstest/orig/
"/mnt/c/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to ods --outdir ../_ODS *.xl*

#Konverter til OOXML
echo "Konverterer ODS til XLSX"
cd /mnt/c/Regneark_fra_s_drevet/konverteringstest/_ODS/
"/mnt/c/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to xlsx --outdir ../_XLSX *.ods

#Konverter til XLS
echo "Konverterer ODS til XLS"
cd /mnt/c/Regneark_fra_s_drevet/konverteringstest/_ODS/
"/mnt/c/Program Files/LibreOffice/program/soffice.exe" --headless --convert-to xls --outdir ../_XLS *.ods

cd /mnt/c/Regneark_fra_s_drevet/konverteringstest/

echo "Kører SpreadsheetComplexityAnalyzer på originalfiler"
#Run SCA on original files
java -jar "-Dfile.encoding=UTF-8" "/mnt/c/SpreadsheetComplexityAnalyzer/SpreadsheetComplexityAnalyser.jar" /mnt/c/Regneark_fra_s_drevet/konverteringstest/orig/ -x > logorig.xml

echo "Kører SpreadsheetComplexityAnalyzer på konverterede XLSX filer"
#Run SCA on converted files
java -jar "-Dfile.encoding=UTF-8" "/mnt/c/SpreadsheetComplexityAnalyzer/SpreadsheetComplexityAnalyser.jar"  /mnt/c/Regneark_fra_s_drevet/konverteringstest/_XLS/ -x > logorig2.xml

echo "Kører SpreadsheetComplexityAnalyzer på konverterede XLS filer"
#Run SCA on converted files
java -jar "-Dfile.encoding=UTF-8" "/mnt/c/SpreadsheetComplexityAnalyzer/SpreadsheetComplexityAnalyser.jar"  /mnt/c/Regneark_fra_s_drevet/konverteringstest/_XLSX/ -x > logorig3.xml

echo "Kombiner XML filer til een fil"
#Kombiner XML filer til een fil
xml_grep --pretty_print indented --wrap "spreadsheetComplexityAnalyserResults" --descr '' --cond "spreadsheetComplexityAnalyserResult" logorig*.xml > combined.xml