/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package qltv;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.Iterator;
import javax.swing.JComboBox;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author Pham Ngoc Minh
 */
public class Statistical extends javax.swing.JPanel {
    private int choose = 0;
    /**
     * Creates new form Statistical
     */
    public Statistical() {
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        statisticalTable = new javax.swing.JTable();
        employee = new javax.swing.JComboBox<>();
        customer = new javax.swing.JComboBox<>();
        book = new javax.swing.JComboBox<>();
        export = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        timeInput = new javax.swing.JTextField();
        timeStatistical = new javax.swing.JComboBox<>();
        jLabel5 = new javax.swing.JLabel();
        jSeparator1 = new javax.swing.JSeparator();

        setBackground(new java.awt.Color(255, 255, 255));

        statisticalTable.setBackground(new java.awt.Color(255, 255, 255));
        statisticalTable.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        statisticalTable.setRowHeight(26);
        jScrollPane1.setViewportView(statisticalTable);

        employee.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Chọn", "Tên", "Giới tính", "Năm sinh" }));
        employee.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                employeeActionPerformed(evt);
            }
        });

        customer.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Chọn", "Địa chỉ", "Giới tính", "Số lượng sách mượn" }));
        customer.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                customerActionPerformed(evt);
            }
        });

        book.setBackground(new java.awt.Color(255, 255, 255));
        book.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Chọn", "Thể loại", "Tên sách", "Tên tác giả", "Tên NXB" }));
        book.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                bookActionPerformed(evt);
            }
        });

        export.setBackground(new java.awt.Color(255, 255, 255));
        export.setText("Xuất");
        export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(51, 51, 51));
        jLabel1.setText("Thống kê Sách");

        jLabel2.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(51, 51, 51));
        jLabel2.setText("Thống kê Nhân viên");

        jLabel3.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(51, 51, 51));
        jLabel3.setText("Thống kê Người mượn");

        timeInput.setBackground(new java.awt.Color(255, 255, 255));
        timeInput.setFont(new java.awt.Font("Segoe UI", 0, 12)); // NOI18N
        timeInput.setBorder(null);
        timeInput.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                timeInputActionPerformed(evt);
            }
        });

        timeStatistical.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Chọn", "Toàn bộ", "Nhân viên", "Khách hàng" }));
        timeStatistical.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                timeStatisticalActionPerformed(evt);
            }
        });

        jLabel5.setFont(new java.awt.Font("Segoe UI", 1, 14)); // NOI18N
        jLabel5.setForeground(new java.awt.Color(51, 51, 51));
        jLabel5.setText("Thống kê theo thời gian");

        jSeparator1.setBackground(new java.awt.Color(255, 255, 255));
        jSeparator1.setForeground(new java.awt.Color(153, 153, 153));

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(12, 12, 12)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(book, javax.swing.GroupLayout.PREFERRED_SIZE, 133, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(173, 173, 173)
                        .addComponent(employee, javax.swing.GroupLayout.PREFERRED_SIZE, 134, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel1)
                        .addGap(210, 210, 210)
                        .addComponent(jLabel2)))
                .addGap(176, 176, 176)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(customer, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(215, 215, 215)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel5)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(timeStatistical, javax.swing.GroupLayout.PREFERRED_SIZE, 93, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(3, 3, 3)
                                .addComponent(timeInput, javax.swing.GroupLayout.PREFERRED_SIZE, 157, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 157, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(0, 0, Short.MAX_VALUE))
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addGap(0, 0, Short.MAX_VALUE)
                        .addComponent(export, javax.swing.GroupLayout.PREFERRED_SIZE, 66, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jScrollPane1))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(6, 6, 6)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel1)
                    .addComponent(jLabel2, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(jLabel5)
                        .addComponent(jLabel3)))
                .addGap(4, 4, 4)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(timeInput, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(4, 4, 4)
                        .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(2, 2, 2)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(book, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(timeStatistical, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(customer, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addComponent(employee, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(20, 20, 20)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 521, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(export)
                .addContainerGap())
        );
    }// </editor-fold>//GEN-END:initComponents

    private void bookActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bookActionPerformed
        // TODO add your handling code here:
        ConnectDB connectDB = new ConnectDB();
        Connection con = connectDB.getConnect();
        JComboBox<String> combo = (JComboBox<String>) evt.getSource();
        String selected = (String) combo.getSelectedItem();
        if (null != selected) {
            switch (selected) {
                case "Thể loại":
                    choose = 1;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Thể loại", "Số lượng"}, 0));
                    String sql = "select Kind, count(*) from qltv.infor_book group by Kind";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Tên sách":
                    choose = 2;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Tên sách", "Số lượng"}, 0));
                    String sql1 = "select Book_name, count(*) from qltv.infor_book group by Book_name";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql1);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Tên tác giả":
                    choose = 3;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Tên tác giả", "Số lượng"}, 0));
                    String sql2 = "select Author, count(*) from qltv.infor_book group by Author";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql2);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Tên NXB":
                    choose = 4;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Tên Nhà xuất bản", "Số lượng"}, 0));
                    String sql3 = "select Publisher, count(*) from qltv.infor_book group by Publisher";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql3);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                default:
                    break;
            }
        }
    }//GEN-LAST:event_bookActionPerformed

    private void exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportActionPerformed
        // TODO add your handling code here:
        try {
            String path = "";
            switch(choose) {
                case 1:
                    path = "";
                    break;
                case 2:
                    path = "";
                case 3:
                    path = "";
                    break;
                case 4:
                    path = "";
                case 5:
                    path = "";
                    break;
                case 6:
                    path = "";
                case 7:
                    path = "";
                    break;
                case 8:
                    path = "";
                case 9:
                    path = "";
                    break;
                case 10:
                    path = "";
                    break;
            }
            FileInputStream fis = new FileInputStream(path);
            XWPFDocument xdoc = new XWPFDocument(OPCPackage.open(fis));
            Iterator bodyElementIterator = xdoc.getBodyElementsIterator();
            while (bodyElementIterator.hasNext()) {
                IBodyElement element = (IBodyElement) bodyElementIterator.next();
                if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                    java.util.List<XWPFTable> tableList = element.getBody().getTables();
                    for (XWPFTable table : tableList) {
                        setDefaultTable(table);
                        for (int i = 1; i < table.getRows().size(); i++) {
                            for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
                                removeParagraphs(table.getRow(i).getCell(j));
                                XWPFParagraph paragraph = table.getRow(i).getCell(j).addParagraph();
                                paragraph.createRun().setText(statisticalTable.getValueAt(i - 1, j).toString());
                            }

                        }
                        addRowData(table, table.getRows().size(), statisticalTable);
                    }
                }
            }
            OutputStream out = new FileOutputStream("C:\\Users\\Pham Ngoc Minh\\Desktop\\TestWord.docx");
            xdoc.write(out);
            out.close();

        } catch (IOException | InvalidFormatException ex) {
        }
        int dialogResult = JOptionPane.showConfirmDialog(null, "File đã tạo thành công!\nBạn có muốn mở file?");
        if (dialogResult == JOptionPane.YES_OPTION) {
            if (Desktop.isDesktopSupported()) {
                try {
                    File myFile = new File("C:\\Users\\Pham Ngoc Minh\\Desktop\\TestWord.docx");
                    Desktop.getDesktop().open(myFile);
                } catch (IOException ex) {
                    // no application registered for PDFs
                }
            }
        } else {
        }
    }//GEN-LAST:event_exportActionPerformed

    private void timeInputActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_timeInputActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_timeInputActionPerformed

    private void timeStatisticalActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_timeStatisticalActionPerformed
        // TODO add your handling code here:
        ConnectDB connectDB = new ConnectDB();
        Connection con = connectDB.getConnect();
        JComboBox<String> combo = (JComboBox<String>) evt.getSource();
        String selected = (String) combo.getSelectedItem();
        if (selected != null) {
            switch (selected) {
                case "Toàn bộ":
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Số phiếu lập", "Tổng số tiền đặt cọc", "Tổng số tiền phạt"}, 0));
                    String sql = "select count(*), sum(Deposits), sum(Mulct) from qltv.loan_payment, qltv.loan_payment_detail where qltv.loan_payment.idloan = qltv.loan_payment_detail.idloan and Borrow_day like '" + timeInput.getText() + "%'";
                    try {
                        PreparedStatement pst = con.prepareStatement(sql);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2), rs.getString(3)});
                        }
                    } catch (Exception e) {
                    }
                    break;
                case "Nhân viên":
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Mã nhân viên","Số phiếu lập", "Tổng số tiền đặt cọc", "Tổng số tiền phạt"}, 0));
                    String sql2 = "select idEmployee, count(*) , sum(Deposits), sum(Mulct) from qltv.loan_payment, qltv.loan_payment_detail \n" +
"where qltv.loan_payment.idloan = qltv.loan_payment_detail.idloan and Borrow_day like'" + timeInput.getText() + "%'group by idEmployee";
                    try {
                        PreparedStatement pst = con.prepareStatement(sql2);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2), rs.getString(3), rs.getString(4)});
                        }
                    } catch (Exception e) {
                    }
                    break;
            }

        }

    }//GEN-LAST:event_timeStatisticalActionPerformed

    private void employeeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_employeeActionPerformed
        // TODO add your handling code here:
        ConnectDB connectDB = new ConnectDB();
        Connection con = connectDB.getConnect();
        
        JComboBox<String> combo = (JComboBox<String>) evt.getSource();
        String selected = (String) combo.getSelectedItem();
        if (selected != null) {
            switch (selected) {
                case "Tên":
                    choose = 5;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Tên Nhân viên", "Số lượng"}, 0));
                    String sql = "select name, count(*) from qltv.employee group by name";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Giới tính":
                    choose = 6;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Giới tính", "Số lượng"}, 0));
                    String sql2 = "select Sex, count(*) from qltv.employee group by Sex";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql2);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Năm sinh":
                    choose = 7;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Ngày sinh", "Số lượng"}, 0));
                    String sql3 = "select BOD, count(*) from qltv.employee group by BOD";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql3);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
            }
        }
    }//GEN-LAST:event_employeeActionPerformed

    private void customerActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_customerActionPerformed
        // TODO add your handling code here:
        ConnectDB connectDB = new ConnectDB();
        Connection con = connectDB.getConnect();
        
        JComboBox<String> combo = (JComboBox<String>) evt.getSource();
        String selected = (String) combo.getSelectedItem();
        if (selected != null) {
            switch (selected) {
                case "Địa chỉ":
                    choose = 8;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Địa chỉ", "Số lượng"}, 0));
                    String sql = "select Address, count(*) from qltv.customer group by Address";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Giới tính":
                    choose = 9;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Giới tính", "Số lượng"}, 0));
                    String sql2 = "select Sex, count(*) from qltv.customer group by Sex";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql2);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                case "Số lượng sách mượn":
                    choose = 10;
                    statisticalTable.setModel(new DefaultTableModel(new Object[]{"Tên Người mượn", "Số lượng"}, 0));
                    String sql3 = "select name, quantity from qltv.customer c inner join (select idcustomer, quantity from qltv.loan_payment, qltv.loan_payment_detail  where qltv.loan_payment.idloan = qltv.loan_payment_detail.idloan group by idcustomer) temp on c.idcustomer = temp.idcustomer";
                    try {
                        con.setAutoCommit(false);
                        PreparedStatement pst = con.prepareStatement(sql3);
                        ResultSet rs = pst.executeQuery();
                        while (rs.next()) {
                            ((DefaultTableModel) statisticalTable.getModel()).addRow(new Object[]{rs.getString(1), rs.getString(2)});
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
            }
        }
    }//GEN-LAST:event_customerActionPerformed

    private void addRowData(XWPFTable table, int lastRowPosition, JTable dataTable) {
        for (int i = lastRowPosition - 1; i < dataTable.getRowCount(); i++) {
            XWPFTableRow newRow = table.createRow();
            for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
                newRow.getCell(j).setText(dataTable.getValueAt(i, j).toString());
            }

        }
    }

    private static void removeParagraphs(XWPFTableCell tableCell) {
        int count = tableCell.getParagraphs().size();
        for (int i = 0; i < count; i++) {
            tableCell.removeParagraph(i);
        }
    }

    private void setDefaultTable(XWPFTable table) {
        for (int i = 1; i < table.getRows().size(); i++) {
            table.removeRow(1);
        }
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> book;
    private javax.swing.JComboBox<String> customer;
    private javax.swing.JComboBox<String> employee;
    private javax.swing.JButton export;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JTable statisticalTable;
    private javax.swing.JTextField timeInput;
    private javax.swing.JComboBox<String> timeStatistical;
    // End of variables declaration//GEN-END:variables
}
