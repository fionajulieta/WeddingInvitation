//package org.example;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.util.Iterator;
//
//public class WeddingInvitationGenerator {
//
//    public static void main(String[] args) {
//        String inputFile = "guests.xlsx";
//        String outputFile = "master_invitations.xlsx";
//
//        try (FileInputStream fis = new FileInputStream(new File(inputFile));
//             Workbook inputWorkbook = new XSSFWorkbook(fis);
//             Workbook outputWorkbook = new XSSFWorkbook()) {
//
//            Sheet inputSheet = inputWorkbook.getSheetAt(0);
//            Sheet outputSheet = outputWorkbook.createSheet("Invitations");
//
//            Iterator<Row> rowIterator = inputSheet.iterator();
//            int outputRowNum = 0;
//
//            // Write header
//            Row headerRow = outputSheet.createRow(outputRowNum++);
//            headerRow.createCell(0).setCellValue("Name");
//            headerRow.createCell(1).setCellValue("Message");
//
//            // Skip header in input
//            if (rowIterator.hasNext()) rowIterator.next();
//
//            while (rowIterator.hasNext()) {
//                Row row = rowIterator.next();
//                String name = row.getCell(0).getStringCellValue();
//                String link = row.getCell(1).getStringCellValue();
//
//                String message = "Dear " + name + "!\n\n" +
//                        "I hope this message finds you well. We are delighted to share some special news - We are getting married!\n\n" +
//                        "We'd love for you to join us at our wedding reception:\n" +
//                        "Date: Saturday, 14 June 2025\n" +
//                        "Time: 18:30 - end\n" +
//                        "Venue: PIK Office Ballroom\n\n" +
//                        "Please feel free to check out all the event details in our digital wedding invitation below and send us a RSVP when you have a moment.\n" +
//                        link + "\n\n" +
//                        "Your presence would mean so much to us, and we really hope you can be part of this unforgettable day.\n\n" +
//                        "Much love,\n" +
//                        "Steven & Fiona";
//
//
//                Row outRow = outputSheet.createRow(outputRowNum++);
//                outRow.createCell(0).setCellValue(name);
//                outRow.createCell(1).setCellValue(message);
//            }
//
//            // Save output
//            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
//                outputWorkbook.write(fos);
//            }
//
//            System.out.println("✅ Master Excel created: " + outputFile);
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//}
package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class WeddingInvitationGenerator {

    public static void main(String[] args) {
        String inputFile = "guests.xlsx";
        String outputFile = "master_invitations.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(inputFile));
             Workbook inputWorkbook = new XSSFWorkbook(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = inputWorkbook.getSheetAt(0);
            Sheet outputSheet = outputWorkbook.createSheet("Invitations");

            Iterator<Row> rowIterator = inputSheet.iterator();
            int outputRowNum = 0;

            // Write header
            Row headerRow = outputSheet.createRow(outputRowNum++);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Versi English");
            headerRow.createCell(2).setCellValue("Versi Indonesia");

            // Skip header in input
            if (rowIterator.hasNext()) rowIterator.next();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String name = row.getCell(0).getStringCellValue();
                String link = row.getCell(1).getStringCellValue();

                String message = "Dear " + name + "!\n\n" +
                        "I hope this message finds you well. We are delighted to share some special news - We are getting married!\n\n" +
                        "We'd love for you to join us at our wedding reception:\n" +
                        "Date: Saturday, 14 June 2025\n" +
                        "Time: 18:30 - end\n" +
                        "Venue: PIK Office Ballroom\n\n" +
                        "Please feel free to check out all the event details in our digital wedding invitation below and send us a RSVP when you have a moment.\n" +
                        link + "\n\n" +
                        "Your presence would mean so much to us, and we really hope you can be part of this unforgettable day.\n\n" +
                        "Much love,\n" +
                        "Steven & Fiona";

                String shortMessage = "Kepada Bpk./Ibu/saudara/i " + name + "!\n\n" +
                        "Dengan sukacita, kami mengundang Bapak/Ibu/Saudara/i untuk menghadiri acara pernikahan kami:\n" +
                        "Steven & Fiona\n" +
                        "\n" +
                        "Seluruh detail acara dapat Anda lihat melalui undangan digital kami di bawah ini. Jika berkenan, kami akan sangat senang menerima konfirmasi kehadiran melalui form RSVP.\n" +
                        link + "\n\n" +
                        "Merupakan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i dapat hadir dan menjadi bagian dari momen berbahagia ini.\n" +
                        "\n" +
                        "Mohon maaf, undangan ini hanya dapat dibagikan melalui pesan ini.\n" +
                        "Atas perhatian dan kehadirannya, kami ucapkan terima kasih.";

                Row outRow = outputSheet.createRow(outputRowNum++);
                outRow.createCell(0).setCellValue(name);
                outRow.createCell(1).setCellValue(message);
                outRow.createCell(2).setCellValue(shortMessage);
            }

            // Autosize columns for better readability
            outputSheet.autoSizeColumn(0);
            outputSheet.autoSizeColumn(1);
            outputSheet.autoSizeColumn(2);

            // Save output
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }

            System.out.println("✅ Master Excel created: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
