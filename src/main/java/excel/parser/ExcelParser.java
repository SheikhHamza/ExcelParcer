package excel.parser;

import org.apache.commons.lang.StringEscapeUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.WordUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.OptionalInt;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

public class ExcelParser{

    public static final int ZERO = 0;

    public List<Fields> processExcelFile(String filePath) {

        File myFile = new File(filePath);
        List<Fields> fields = new ArrayList <>();

        try {
            FileInputStream fis = new FileInputStream(myFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet firstSheet = workbook.getSheetAt(ZERO);

            Row row = firstSheet.getRow(ZERO);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();

            for (short cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                String refinedColumnName;
                String originalColumnName = cell.getStringCellValue();
                Pattern pattern = Pattern.compile("^(\\d+.*|-\\d+.*)");
                Matcher matcher = pattern.matcher(originalColumnName);

                refinedColumnName = originalColumnName.toUpperCase()
                        .replaceAll("[^\\w]+", String.valueOf("_"));

                refinedColumnName = StringUtils.remove(WordUtils.capitalizeFully(refinedColumnName,"_".toCharArray()), "_");
                char[] colName = refinedColumnName.toCharArray();
                colName[0] = Character.toLowerCase(colName[0]);

                refinedColumnName = new String(colName);

                // extract number from string (if a string starts with number) and then append at the end of string
                if(matcher.matches()){
                    final String forLambda = refinedColumnName;
                    OptionalInt firstLetterIndex = IntStream.range(0, forLambda.length())
                            .filter(i -> Character.isLetter(forLambda.charAt(i)))
                            .findFirst();
                    if(firstLetterIndex.isPresent()) {
                        String numbers = refinedColumnName.substring(0, firstLetterIndex.getAsInt());
                        refinedColumnName = refinedColumnName.substring(firstLetterIndex.getAsInt()) + numbers;
                    }
                }

                fields.add(new Fields(refinedColumnName,originalColumnName));
            }

            firstSheet.removeRow(row);
        }
        catch (Exception e) {
            throw new RuntimeException("File Processing Error " + e.getMessage());
        }
        return fields;
    }

    private void createJavaClass(String packageName,String className,Boolean isDto,
                                 List<Fields> fields, String javaOutputFilePath) throws IOException{
        PrintWriter out = new PrintWriter(new BufferedWriter(new FileWriter(javaOutputFilePath + (isDto? "Dto":"") + ".java")));
        System.out.println("generating class..");
        out.println("package "+packageName+";" + "\n\n");
        out.println("public class " + className + (isDto? "Dto":"") + " {" + "\n" );

        for (Fields x: fields) {
            if(isDto) {
                out.println("\t@Value(\"" + x.getRefinedName() + "\")");
            }
            out.println("\tprivate String " + x.getRefinedName() +  ";");
        }

        for (Fields x:fields) {
            String tempField=StringEscapeUtils.escapeJava(x.getRefinedName()).substring(0, 1).toUpperCase()+StringEscapeUtils.escapeJava(x.getRefinedName()).substring(1);

            //getter method
            out.println("");
            out.println("\tpublic String get"+tempField+ "(){");
            out.println("\t\t return this."+x.getRefinedName()+";");
            out.println("\t}");
            //setter method
            out.println("\tpublic void set"+tempField+"(String "+ StringEscapeUtils.escapeJava(x.getRefinedName())+"){");
            out.println("\t\t this."+StringEscapeUtils.escapeJava(x.getRefinedName())+" = "+ StringEscapeUtils.escapeJava(x.getRefinedName())+";");
            out.println("\t}");
        }

        out.println("");
        out.println("}");
        out.close();
    }

    public static void main(String args[]){

        // Sample execution
        ExcelParser parser = new ExcelParser();
        List<Fields> fields = parser.processExcelFile("C:\\Users\\VenD\\Desktop\\TempDN\\MasterDataFields.xlsx");
        System.out.println(fields.get(1).toString());
        try {
            parser.createJavaClass("excel.parser","MasterDataFields",true,fields,"C:\\Users\\VenD\\Desktop\\TempDN\\MasterDataFields");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
