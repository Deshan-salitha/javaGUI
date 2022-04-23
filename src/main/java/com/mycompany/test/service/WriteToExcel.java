/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.test.service;

import com.mycompany.test.model.User;
import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 *
 * @author Harry Bob
 */
public class WriteToExcel {
    public static void addRowToExcel(List<User> user) {
        
        //blank wookbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //create blank sheet
        XSSFSheet sheet = workbook.createSheet();
//        XSSFSheet sheet = new XSSFSheet();
        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        for (int i = 0; i < user.size(); i++) {
            System.out.println(i);
            System.out.println(user.get(i).getId());
            System.err.println(user);
            System.out.println(user.get(0).toString());
            data.put(Integer.toString(i), new Object[]{user.get(i).getId(), user.get(i).getName(), user.get(i).getEmail()});
        }

        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                }
            }
        }
        try {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("Users.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Users.xlsx written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
//            try {
//
//            FileWriter writer = new FileWriter("User.txt", true);
//            writer.write(u.getId());
//            writer.write(u.getName());
//            writer.write(u.getEmail());
//            JOptionPane.showMessageDialog(rootPane, "Success");
//
//            } catch (Exception e) {
//               JOptionPane.showMessageDialog(rootPane, "Error");
//            }

    }
}
