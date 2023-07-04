package org.example;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

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


    CollectorFrame() {


        JFrame jFrame = new JFrame("Excel Collector");
        jFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        jFrame.setSize(300, 100);
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
                    String firstLine = "CONTACT ANGLE" + "\n" + "LOG BOOK";
                    String secondLine = "ANALYSIS / PROCESS";
                    String thirdLine = "EQUIPMENT NAME,DIVISION,DATE,TIME,OPERATION,USER,PURPOSE OF OPERATION,PROJECT CODE,USAGE DURATION,USAGE MODE,INSTITUTION NAME,INSTITUTION TYPE,PERSON NAME&SURNAME,PERSON TITLE,PERSON POSITON";
                    writer.write(firstLine);
                    writer.write("\n" + secondLine);
                    writer.write("\n" + thirdLine);
                    outerLoop:
                    for (File file : files) {
                        if (file.getName().endsWith(".xlsx")) {
                            path = file.getPath();
                            String fileName = file.getName().substring(0, file.getName().length() - 5);

                           if (fileName.contains(",")) {
                                for (int fn = 0; fn <= fileName.length(); fn++) {
                                    fileName = fileName.replace(",", "_");
                                }
                            }

                            System.out.println("----------------COLLECTOR STARTING A NEW FİLE----------------");

                            System.out.println("FİLE PATH :" + path);

                            try {
                                FileInputStream inputStream = new FileInputStream(path);
                                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                                XSSFSheet sheet = workbook.getSheet("Chart");
                                XSSFCell cell;
                                XSSFCell cellCheck;
                                kacDosyaOkudum++;

                               rowLoop: for (int i = 3; i <= sheet.getLastRowNum(); i++) {
                                    String value = "";
                                    for (int j = 0; j <= 14; j++) {

                                        //CELL TANIMLANDI,J KAÇ İSE O HÜCREDEYİZ
                                        cell = sheet.getRow(i).getCell(j);
                                        if (cell != null) {
                                            if (cell.getCellType() != CellType.BLANK) {
                                                if (j == 1) {
                                                    if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                                                        Date date = cell.getDateCellValue();
                                                        DateFormat timeFormat = new SimpleDateFormat("dd/MM/yy");
                                                        String timeString = timeFormat.format(date);
                                                        value += timeString + ",";

                                                    } else {
                                                        value += cell + ",";
                                                    }
                                                } else if (j == 2) {
                                                    //System.out.println("CELL TYPE : "+cell.getCellType());
                                                    if (cell.getCellType() == CellType.STRING) {
                                                        value += cell.getStringCellValue() + ",";
                                                        continue;
                                                    } else if (cell.getCellType() == CellType.NUMERIC) {
                                                        String timeString = cell.getLocalDateTimeCellValue().toString();
                                                        String[] times = timeString.split("T");
                                                        String startTimeString = times[0];
                                                        String endTimeString = times[1];
                                                        value += endTimeString + ",";
                                                        continue;
                                                    } else {
                                                        if (cell.getCellType() == CellType.BLANK) {
                                                            value += ",";
                                                            continue;
                                                        }
                                                    }
                                                } else {
                                                    value += getCellValue(cell) + ",";
                                                }

                                            }
                                            else {
                                                value+=",";


                                            }

                                        }
                                        //ard arda 4 cell boşmu
                                        else{
                                            boolean allCellsEmpty = true;
                                            for (Cell cell1 : sheet.getRow(i)) {
                                                if (cell1 != null && cell1.getCellType() != CellType.BLANK) {
                                                    allCellsEmpty = false;
                                                    break;
                                                }
                                            }
                                            if (allCellsEmpty) {
                                                continue outerLoop;
                                            } else {
                                                value+=",";

                                            }
                                        }


                                    }
                                    System.out.println("\n " + fileName + "," + value + ",");
                                    writer.write("\n " + fileName + "," + value + ",");
                                }

                            } catch (Exception exception) {
                                exception.printStackTrace();
                            }

                        }

                    }
                    System.out.println(kacDosyaOkudum + " ADET DOSYA OKUNDU !");
                    writer.close();
                } catch (IOException exception) {
                    exception.printStackTrace();
                }


            }

        });


        jPanel.add(button);
        jFrame.add(jPanel);

        jFrame.setResizable(false);
        jFrame.setVisible(true);


    }


    public String getCellValue(XSSFCell cell) {


        switch (cell.getCellType()) {

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
