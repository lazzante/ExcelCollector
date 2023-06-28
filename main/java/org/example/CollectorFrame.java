package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static java.util.Calendar.DATE;

public class CollectorFrame extends JFrame {
    JButton button;
    String path;



    CollectorFrame(){



        JFrame jFrame = new JFrame("Excel Collector");
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        jFrame.setSize(300,100);
        JPanel jPanel = new JPanel();

        button = new JButton("Start Collector");
        button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String systemPath = System.getProperty("user.dir");
                File filepath = new File(systemPath);
                File[] files = filepath.listFiles();
                int kacDosyaOkudum = 0;


                try {
                    BufferedWriter writer = new BufferedWriter(new FileWriter("CollectorsOutput.csv"));
                    String firstLine = "CONTACT ANGLE"+"\n"+"LOG BOOK";
                    String secondLine = "ANALYSIS / PROCESS";
                    String thirdLine = "EQUIPMENT NAME,DIVISION,DATE,TIME,OPERATION,USER,PURPOSE OF OPERATION,PROJECT CODE,USAGE DURATION,USAGE MODE,INSTITUTION NAME,INSTITUTION TYPE,PERSON NAME&SURNAME,PERSON TITLE,PERSON POSITON";
                    writer.write(firstLine);
                    writer.write("\n"+secondLine);
                    writer.write("\n"+thirdLine);

                    List<String> cellsList = new ArrayList<String>();

                    for (File file : files) {
                        if (file.getName().endsWith(".xlsx")) {
                            path = file.getPath();
                            String fileName =file.getName().substring(0,file.getName().length()-5);

                            System.out.println("----------------COLLECTOR STARTING A NEW FİLE----------------");

                            System.out.println("FİLE PATH :" + path);

                            try {
                                FileInputStream inputStream = new FileInputStream(path);

                                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

                                XSSFSheet sheet = workbook.getSheet("Chart");

                                XSSFCell cell;



                                String cellValue="";
                                kacDosyaOkudum++;

                                for(int i=3;i<=sheet.getLastRowNum();i++){
                                    String value="";
                                    System.out.println("\n");
                                    for(int j = 0;j<=sheet.getRow(i).getLastCellNum();j++){

                                        //CELL TANIMLANDI,J KAÇ İSE O HÜCREDEYİZ
                                        cell = sheet.getRow(i).getCell(j);

                                        if(sheet.getRow(i)==null){
                                            System.out.println("BREAK ÇALIŞTI");
                                            break;
                                        }

                                        else if(cell==null){
                                            /*cell = sheet.getRow(i).getCell(j,Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                                            cellValue="";*/
                                            value+=" ";
                                        }

                                        else{
                                            if(j==1){
                                                if(cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)){
                                                    Date date= cell.getDateCellValue();
                                                    DateFormat timeFormat = new SimpleDateFormat("dd/MM/yy");
                                                    String timeString = timeFormat.format(date);
                                                    value+=timeString+",";

                                                }else{
                                                   value+=cell+",";
                                                }
                                            }
                                            else if(j==2){
                                                System.out.println("CELL TYPE : "+cell.getCellType());
                                                //EN SON BURADA KALDIM GELENİN SAAT OLUP OLMADIĞINI KONTROL ETMEM LAZIM
                                                if(cell.getCellType()==CellType.STRING){
                                                    value+=cell.getStringCellValue()+",";
                                                    continue;
                                                }
                                                else if (cell.getCellType()==CellType.NUMERIC) {
                                                    String timeString = cell.getLocalDateTimeCellValue().toString();
                                                    String[] times = timeString.split("T");
                                                    String startTimeString = times[0];
                                                    String endTimeString = times[1];
                                                    value+=endTimeString+",";
                                                    continue;
                                                }
                                                else{
                                                    if(cell.getCellType()==CellType.BLANK){
                                                        value+=" ";
                                                        continue;
                                                    }
                                                }
                                            }
                                            else {
                                                value+=cell.toString()+",";
                                            }


                                        }

                                        if(sheet.getRow(i).getCell(0)==null&&sheet.getRow(i).getCell(1)==null&&sheet.getRow(i).getCell(2)==null&&sheet.getRow(i).getCell(3)==null&&
                                                sheet.getRow(i).getCell(4)==null&&sheet.getRow(i).getCell(5)==null&&sheet.getRow(i).getCell(6)==null&&sheet.getRow(i).getCell(7)==null&&
                                                sheet.getRow(i).getCell(8)==null&&sheet.getRow(i).getCell(9)==null&&sheet.getRow(i).getCell(10)==null&&sheet.getRow(i).getCell(11)==null&&
                                                sheet.getRow(i).getCell(12)==null&&sheet.getRow(i).getCell(13)==null&&sheet.getRow(i).getCell(14)==null&&sheet.getRow(i).getCell(15)==null){
                                            workbook.close();
                                            inputStream.close();
                                            break;
                                        }

                                    }
                                    System.out.println("\n "+i+" "+value);






                                }

                            } catch (Exception exception) {
                                System.out.println("HATA MESAJI(SATIR 147) : "+exception.getMessage());
                                exception.printStackTrace();
                            }

                        }

                    }
                    System.out.println(kacDosyaOkudum+" ADET DOSYA OKUNDU !");
                    writer.close();
                }catch (IOException exception){
                    exception.printStackTrace();
                }


            }

        });


        jPanel.add(button);
        jFrame.add(jPanel);

        jFrame.setResizable(false);
        jFrame.setVisible(true );


    }








    public String getCellValue(XSSFCell cell){


        switch (cell.getCellType()){

            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case STRING:
                return cell.getStringCellValue();

            default:
                return String.valueOf(cell.getStringCellValue());

        }
    }
}
