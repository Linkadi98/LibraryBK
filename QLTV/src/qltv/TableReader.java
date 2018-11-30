/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package qltv;

import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.temporal.ChronoUnit;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;



public class TableReader {

    public static long betweenDates(Date firstDate, Date secondDate) throws IOException {
        return ChronoUnit.DAYS.between(firstDate.toInstant(), secondDate.toInstant());
    }
    
    public static void main(String[] args) throws IOException {
        
        DateFormat format = new SimpleDateFormat("yyyy-MM-dd");
        try {
            Date date1 = format.parse("2018-8-1");
            Date date2 = format.parse("2019-10-1");
            System.out.println(betweenDates(date1, date2));
        } catch (ParseException ex) {
            Logger.getLogger(TableReader.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
} // ông thử chạy xem 

