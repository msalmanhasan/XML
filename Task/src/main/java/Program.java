import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;
import java.awt.*;
import java.io.*;

public class Program {

    //Path for Excel file
    public static final String XLSM_FILE_PATH = "src/main/resources/Amtliche Fassung des ATC-Index 2019.xlsm";

    //Path for XML file to be created
    public static final String OUTPUT_PATH = "src/main/resources/Amtliche Fassung des ATC-Index 2019.xml";


    //Main Program
    public static void main(String[] args){

        Workbook workbook = null;
        PrintWriter out = null;

        try {
            workbook = WorkbookFactory.create(new File(XLSM_FILE_PATH));
            out = new PrintWriter(new BufferedWriter(new FileWriter(OUTPUT_PATH)));

        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        System.out.println("Converting third sheet (amtlicher Index alphabet. 2019) to XML format...");

        Sheet sheet = workbook.getSheet("amtlicher Index alphabet. 2019");

        out.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        out.println("\n\n<amtlicher-Index-alphabet-2019>");

        //loop through records and print to file
        boolean firstRow = true;                    //Skip the first row
        for (Row row : sheet) {
            if (firstRow == true) {
                firstRow = false;
                continue;
            }
            out.println();
            out.println("\t<Record>");
            out.println(formatElement("\t\t", "ATC-CODE", formatCell(row.getCell(0))));
            out.println(formatElement("\t\t", "BEDEUTUNG", formatCell(row.getCell(2))));
            out.println(formatElement("\t\t", "DDD-INFO", formatCell(row.getCell(4))));
            out.println("\t</Record>");
        }

        out.write("\n\n</amtlicher-Index-alphabet-2019>");
        out.flush();
        out.close();

        System.out.println("XML file has been created");

        System.out.println("Now reading the XML file and showing in GUI");

        readAndShowXml();



    }


    //Method that reads XML file and displays the content in GUI
    private static void readAndShowXml() {
        BufferedReader bufferedReader = null;
        StringBuilder sb = new StringBuilder();
        String line = null;

        try {
            bufferedReader = new BufferedReader(new FileReader(OUTPUT_PATH));
            while ((line =bufferedReader.readLine())!=null){

                sb.append(line + '\n');
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }

        //System.out.println("Created the string of all text in XML file.");

        JFrame frame = new JFrame("XML GUI");
        JTextArea area = new JTextArea(40,80);

        JScrollPane jScrollPane = new JScrollPane(area,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);

        area.setText(sb.toString());


        frame.setLayout(new FlowLayout());
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.add(jScrollPane);
        frame.setSize(1000,1000);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);

    }

    //Returns the value from cell
    private static String formatCell(Cell cell)
    {
        if (cell == null) {
            return "";
        }
        switch(cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return Boolean.toString(cell.getBooleanCellValue());
            case Cell.CELL_TYPE_ERROR:
                return "*error*";
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                return "<unknown value>";
        }
    }


    //Method to format XML element
    private static String formatElement(String prefix, String tag, String value) {
        StringBuilder sb = new StringBuilder(prefix);
        sb.append("<");
        sb.append(tag);
        if (value != null && value.length() > 0) {
            sb.append(">");
            sb.append(value.replace("\n",""));    //removes the newline character from the value string
            sb.append("</");
            sb.append(tag);
            sb.append(">");
        } else {
            sb.append("/>");
        }
        return sb.toString();
    }

}


