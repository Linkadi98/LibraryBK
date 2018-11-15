/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package qltv;

import java.awt.event.KeyEvent;
import java.awt.Desktop;
import java.awt.Font;
import javax.swing.*;
import java.awt.HeadlessException;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPopupMenu;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
public class Detail extends javax.swing.JPanel {

    private int clicked = 0;

    /**
     * Creates new form Detail
     */
    public Detail() {
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

        buttonGroup1 = new javax.swing.ButtonGroup();
        jScrollPane1 = new javax.swing.JScrollPane();
        loanPaymentTable = new javax.swing.JTable();
        jScrollPane2 = new javax.swing.JScrollPane();
        detailTable = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        checkDetailTable = new javax.swing.JCheckBox();
        checkLoanTabel = new javax.swing.JCheckBox();
        jToolBar1 = new javax.swing.JToolBar();
        chooseFile = new javax.swing.JButton();
        save = new javax.swing.JButton();
        jToolBar2 = new javax.swing.JToolBar();
        insertData = new javax.swing.JButton();
        showData = new javax.swing.JButton();
        clearAll = new javax.swing.JButton();

        loanPaymentTable.setModel(new javax.swing.table.DefaultTableModel(
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
        loanPaymentTable.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                loanPaymentTableMouseClicked(evt);
            }
            public void mouseReleased(java.awt.event.MouseEvent evt) {
                loanPaymentTableMouseReleased(evt);
            }
        });
        loanPaymentTable.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyPressed(java.awt.event.KeyEvent evt) {
                loanPaymentTableKeyPressed(evt);
            }
        });
        jScrollPane1.setViewportView(loanPaymentTable);

        detailTable.setModel(new javax.swing.table.DefaultTableModel(
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
        detailTable.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                detailTableMouseClicked(evt);
            }
        });
        jScrollPane2.setViewportView(detailTable);

        jLabel1.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("Hoá đơn mượn trả");

        jLabel2.setFont(new java.awt.Font("Arial", 1, 14)); // NOI18N
        jLabel2.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel2.setText("Chi tiết hoá đơn");

        checkDetailTable.setEnabled(false);
        checkDetailTable.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                checkDetailTableActionPerformed(evt);
            }
        });

        checkLoanTabel.setEnabled(false);

        jToolBar1.setRollover(true);

        chooseFile.setIcon(new javax.swing.ImageIcon(getClass().getResource("/qltv/icons8_Microsoft_Excel_25px_3.png"))); // NOI18N
        chooseFile.setToolTipText("Chọn file");
        chooseFile.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        chooseFile.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chooseFileActionPerformed(evt);
            }
        });
        jToolBar1.add(chooseFile);

        save.setForeground(new java.awt.Color(51, 51, 51));
        save.setIcon(new javax.swing.ImageIcon(getClass().getResource("/qltv/icons8_Microsoft_Word_25px.png"))); // NOI18N
        save.setToolTipText("Xuất file");
        save.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        save.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saveActionPerformed(evt);
            }
        });
        jToolBar1.add(save);

        jToolBar2.setRollover(true);

        insertData.setIcon(new javax.swing.ImageIcon(getClass().getResource("/qltv/icons8_Add_Database_25px.png"))); // NOI18N
        insertData.setToolTipText("Thêm vào database");
        insertData.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                insertDataActionPerformed(evt);
            }
        });
        jToolBar2.add(insertData);

        showData.setIcon(new javax.swing.ImageIcon(getClass().getResource("/qltv/icons8_Database_View_25px.png"))); // NOI18N
        showData.setToolTipText("Hiển thị database");
        showData.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        showData.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                showDataActionPerformed(evt);
            }
        });
        jToolBar2.add(showData);

        clearAll.setIcon(new javax.swing.ImageIcon(getClass().getResource("/qltv/icons8_Delete_Document_25px.png"))); // NOI18N
        clearAll.setToolTipText("Xoá bảng");
        clearAll.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                clearAllActionPerformed(evt);
            }
        });
        jToolBar2.add(clearAll);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 1064, Short.MAX_VALUE)
                    .addComponent(jScrollPane1)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                        .addGap(16, 16, 16)
                        .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(checkLoanTabel))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(11, 11, 11)
                        .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGap(28, 28, 28)
                        .addComponent(checkDetailTable))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jToolBar1, javax.swing.GroupLayout.PREFERRED_SIZE, 82, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jToolBar2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jToolBar1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jToolBar2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(7, 7, 7)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(checkLoanTabel)
                    .addComponent(jLabel1, javax.swing.GroupLayout.Alignment.TRAILING))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 250, Short.MAX_VALUE)
                .addGap(31, 31, 31)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(checkDetailTable))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 256, Short.MAX_VALUE)
                .addContainerGap())
        );
    }// </editor-fold>//GEN-END:initComponents

    private void detailTableMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_detailTableMouseClicked
        // TODO add your handling code here:
        checkDetailTable.setSelected(true);
        checkLoanTabel.setSelected(false);
    }//GEN-LAST:event_detailTableMouseClicked

    private void checkDetailTableActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_checkDetailTableActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_checkDetailTableActionPerformed

    private void loanPaymentTableMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_loanPaymentTableMouseClicked
        // TODO add your handling code here:
        checkLoanTabel.setSelected(true);
        checkDetailTable.setSelected(false);
        int row = loanPaymentTable.getSelectedRow();
        System.out.println(row);
        detailTable.setValueAt(loanPaymentTable.getValueAt(row, 0), row, 1);

    }//GEN-LAST:event_loanPaymentTableMouseClicked

    private void loanPaymentTableKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_loanPaymentTableKeyPressed
        // TODO add your handling code here:
        if (evt.getKeyCode() == KeyEvent.VK_ENTER || evt.getKeyCode() == KeyEvent.VK_RIGHT) {
            detailTable.setValueAt(loanPaymentTable.getValueAt(0, 0), 0, 1);
        }
    }//GEN-LAST:event_loanPaymentTableKeyPressed

    private void chooseFileActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chooseFileActionPerformed
        // TODO add your handling code here:
        String path = null;
        JFileChooser fileChooser = new JFileChooser();
        // show ra một bảng chọn file
        int rVal = fileChooser.showOpenDialog(null);
        // nếu nhấn nút ok (tuỳ chọn APPROVE_OPTION)
        if (rVal == JFileChooser.APPROVE_OPTION) {
            String fileName = fileChooser.getSelectedFile().getName();
            String dir = fileChooser.getCurrentDirectory().toString();
            path = dir + "\\" + fileName;
        } // nếu nhấn nút cancel trong bảng
        else if (rVal == JFileChooser.CANCEL_OPTION) {
            return;
        }
        // chỗ này sẽ delete hết các dòng trước khi nhập dữ liệu từ file
        deleteAllRows();
        // vector lưu tên cột
        Vector columns = new Vector();
        try {
            FileInputStream file = new FileInputStream(new File(path));
            // tạo một file excel
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            // tạo một sheet trong excel có số thứ tự là 0
            XSSFSheet sheet = workbook.getSheetAt(0);
            // con trỏ duyệt hàng trong excel
            Iterator<Row> rowIt = sheet.iterator();
            // nếu vẫn còn dòng trong file
            while (rowIt.hasNext()) {
                // tạo một dòng mới
                Row row = rowIt.next();
                // con trỏ trỏ vào các ô trong một dòng
                Iterator<Cell> cellIt = row.cellIterator();
                // nếu là hàng 0
                if (row.getRowNum() == 0) {
                    // add tên các cột vào trong bảng jtable
                    while (cellIt.hasNext()) {
                        Cell cell = cellIt.next();
                        columns.add(cell.getStringCellValue());
                        ((DefaultTableModel) loanPaymentTable.getModel()).setColumnIdentifiers(columns);
                    }
                } else {
                    //vector chứa dữ liệu trong 1 dòng để add vào bảng jtabel
                    Vector<String> rowData = new Vector<String>();
                    // nếu vẫn còn ô tiếp theo
                    while (cellIt.hasNext()) {
                        // lấy cell trong bảng excel
                        Cell cell = cellIt.next();
                        // nếu cell có kiểu dữ liệu là string
                        if (cell.getCellType() == CellType.STRING) {
                            rowData.add(cell.getStringCellValue());
                        } // nếu cell có kiểu dữ liệu là số
                        else if (cell.getCellType() == CellType.NUMERIC) {
                            rowData.add(Double.toString(cell.getNumericCellValue()));
                        }
                    }
                    // add dữ liệu vào trong bảng jtable
                    ((DefaultTableModel) loanPaymentTable.getModel()).addRow(rowData);
                }
            }
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BookTab.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(BookTab.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_chooseFileActionPerformed

    private void saveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saveActionPerformed
        // TODO add your handling code here:
        try {
            FileInputStream fis = new FileInputStream("C:\\Users\\Pham Ngoc Minh\\Desktop\\TestWord.docx");
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
                                paragraph.createRun().setText(loanPaymentTable.getValueAt(i - 1, j).toString());
                            }

                        }
                        addRowData(table, table.getRows().size());
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
    }//GEN-LAST:event_saveActionPerformed

    private void insertDataActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_insertDataActionPerformed
        // TODO add your handling code here:
        ConnectDB connectDB = new ConnectDB();
        Connection connection = connectDB.getConnect();

        //        DefaultTableModel tableModel = (DefaultTableModel) loanPaymentTable.getModel();
        int rows = loanPaymentTable.getRowCount();
        for (int row = 0; row < rows; row++) {
            if (!inserted[row] && !isEmptyRow(row)) {
                String sql = "INSERT INTO qltv. (idBook, Book_name, Author, Publisher, Kind, Cost, Number_books) VALUES (?,?,?,?,?,?,?)";
                try {
                    connection.setAutoCommit(false);
                    PreparedStatement pst = connection.prepareStatement(sql);
                    String idBook = (String) loanPaymentTable.getValueAt(row, 0);
                    String Book_name = (String) loanPaymentTable.getValueAt(row, 1);
                    String Author = (String) loanPaymentTable.getValueAt(row, 2);
                    String Publisher = (String) loanPaymentTable.getValueAt(row, 3);
                    String Kind = (String) loanPaymentTable.getValueAt(row, 4);
                    String Cost = (String) loanPaymentTable.getValueAt(row, 5);
                    String Number_books = (String) loanPaymentTable.getValueAt(row, 6);
                    pst.setString(1, idBook);
                    pst.setString(2, Book_name);
                    pst.setString(3, Author);
                    pst.setString(4, Publisher);
                    pst.setString(5, Kind);
                    pst.setString(6, Cost);
                    pst.setString(7, Number_books);
                    pst.addBatch();
                    pst.executeUpdate();
                    connection.commit();

                } catch (HeadlessException | SQLException ex) {
                    JOptionPane.showMessageDialog(null, ex.getMessage());
                }
                inserted[row] = true;
            }
        }
        JOptionPane.showMessageDialog(null, "Successfully");
    }//GEN-LAST:event_insertDataActionPerformed

    private void showDataActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_showDataActionPerformed
        // TODO add your handling code here:
        String[] defaultColumnNames = {"Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Thể loại", "Giá", "Số lượng"};
        String[][] data = {};
        loanPaymentTable.setModel(new DefaultTableModel(data, defaultColumnNames));
        ConnectDB connectDB = new ConnectDB();
        Connection connection = connectDB.getConnect();
        DefaultTableModel tableModel = (DefaultTableModel) loanPaymentTable.getModel();
        String sql = "SELECT * FROM qltv.";
        PreparedStatement pst;
        int row = 0;
        try {
            pst = connection.prepareStatement(sql);
            ResultSet rs = pst.executeQuery();
            while (rs.next()) {
                tableModel.addRow(new Object[]{rs.getString(1), rs.getString(2),
                    rs.getString(3), rs.getString(4), rs.getString(5), rs.getString(6), rs.getString(7)
                });
                inserted[row] = true;
                row++;
            }
            loanPaymentTable.setModel(tableModel);
        } catch (SQLException ex) {
            Logger.getLogger(BookTab.class.getName()).log(Level.SEVERE, null, ex);
        }
        for (int i = 0; i < 40; i++) {
            tableModel.addRow(new Object[]{null});
        }
//        originalTableModel = (Vector) ((DefaultTableModel) loanPaymentTable.getModel()).getDataVector().clone();
        //        addDocumentListener();
    }//GEN-LAST:event_showDataActionPerformed

    private void clearAllActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_clearAllActionPerformed
        // TODO add your handling code here:
        String[] defaultColumnNames = {"Mã sách", "Tên sách", "Tác giả", "Nhà xuất bản", "Thể loại", "Giá", "Số lượng"};
        String[][] data = {null, null, null, null, null, null, null, null, null, null, null, null, null};
        loanPaymentTable.setModel(new DefaultTableModel(data, defaultColumnNames));
        setDefaultInsertedRows();
    }//GEN-LAST:event_clearAllActionPerformed

    private void loanPaymentTableMouseReleased(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_loanPaymentTableMouseReleased
        // TODO add your handling code here:
        loanPaymentTable.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseReleased(MouseEvent e) {
                int r = loanPaymentTable.rowAtPoint(e.getPoint());
                if (r >= 0 && r < loanPaymentTable.getRowCount()) {
                    loanPaymentTable.setRowSelectionAllowed(true);
                } else {
                    loanPaymentTable.clearSelection();
                }

                int[] rowindex = loanPaymentTable.getSelectedRows();
                for (int i = 0; i < rowindex.length; i++) {
                    int j = rowindex[i];
                    if (j < 0) {
                        return;
                    }
                }
                if (e.isPopupTrigger() && e.getComponent() instanceof JTable) {
                    JPopupMenu popup = popUpLP();
                    popup.show(e.getComponent(), e.getX(), e.getY());
                }
            }
        });
    }//GEN-LAST:event_loanPaymentTableMouseReleased

    private JPopupMenu popUpLP() {
        JPopupMenu popupMenu = new JPopupMenu();
        JMenu deleteMenu = new JMenu("Delete");
        JPopupMenu subPopupMenu = new JPopupMenu();
        JMenuItem deleteFromTb = new JMenuItem("From table");
        JMenuItem deleteFromDb = new JMenuItem("From database");
        JMenu insertMenu = new JMenu("Insert");
        JMenuItem insertAbove = new JMenuItem("Insert Above");
        JMenuItem insertBelow = new JMenuItem("Insert Below");
        JMenuItem update = new JMenuItem("Update");
        Icon icon = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-cancel-16.png");
        Icon deleteDb = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-delete-database-20.png");
        Icon deleteTb = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-delete-table-25.png");
        Icon updateIcon = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-downloading-updates-20.png");
        Icon insertIcon = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-add-row-25.png");
        Icon insertBelowIcon = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-down-arrow-25.png");
        Icon insertAboveIcon = new ImageIcon("C:\\Users\\Pham Ngoc Minh\\Downloads\\Icon\\icons8-long-arrow-up-25.png");
        insertMenu.setIcon(insertIcon);
        insertAbove.setIcon(insertAboveIcon);
        insertBelow.setIcon(insertBelowIcon);
        deleteMenu.setIcon(icon);
        deleteFromDb.setIcon(deleteDb);
        deleteFromTb.setIcon(deleteTb);
        update.setIcon(updateIcon);
        deleteFromDb.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseReleased(MouseEvent e) {
                int[] rows = loanPaymentTable.getSelectedRows();
                Arrays.sort(rows);
                for (int i = 0; i < rows.length; i++) {
                    int row = rows[i];
                    ConnectDB connectDB = new ConnectDB();
                    Connection connection = connectDB.getConnect();

                    DefaultTableModel tableModel = (DefaultTableModel) loanPaymentTable.getModel();
                    String sql = "DELETE FROM qltv. WHERE (idBook = ?)";
                    try {
                        connection.setAutoCommit(false);
                        PreparedStatement pst = connection.prepareStatement(sql);
                        String idBook = (String) tableModel.getValueAt(row, 0);
                        pst.setString(1, idBook);
                        pst.executeUpdate();
                        connection.commit();
                    } catch (HeadlessException | SQLException ex) {
                        JOptionPane.showMessageDialog(null, "Can not delele!\n" + ex.getMessage());
                    }
                    ((DefaultTableModel) loanPaymentTable.getModel()).removeRow(row);
                    for (int j = i + 1; j < rows.length; j++) {
                        rows[j] = rows[j] - 1;
                    }
                }
                JOptionPane.showMessageDialog(null, "Successfully delete");
            }
        });
        deleteFromTb.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseReleased(MouseEvent e) {

                int[] rows = loanPaymentTable.getSelectedRows();
                Arrays.sort(rows);

                for (int i = 0; i < rows.length; i++) {
                    int row = rows[i];
                    ((DefaultTableModel) loanPaymentTable.getModel()).removeRow(row);
                    for (int j = i + 1; j < rows.length; j++) {
                        rows[j] = rows[j] - 1;
                    }
                }
            }
        });
        update.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseReleased(MouseEvent e) {

                ConnectDB connectDB = new ConnectDB();
                Connection connection = connectDB.getConnect();

                DefaultTableModel tableModel = (DefaultTableModel) loanPaymentTable.getModel();
                int[] rows = loanPaymentTable.getSelectedRows();
                Arrays.sort(rows);
                for (int i = 0; i < rows.length; i++) {
                    int row = rows[i];

                    String temp = (String) tableModel.getValueAt(row, 0);
                    String sql = "UPDATE qltv. SET idBook = ?, Book_name = ?, Author = ?, Publisher = ?, Kind = ?, Cost = ?, Number_books = ? WHERE (idBook = ?)";
                    try {
                        connection.setAutoCommit(false);
                        PreparedStatement pst = connection.prepareStatement(sql);
                        String idBook = (String) tableModel.getValueAt(row, 0);
                        String Book_name = (String) tableModel.getValueAt(row, 1);
                        String Author = (String) tableModel.getValueAt(row, 2);
                        String Publisher = (String) tableModel.getValueAt(row, 3);
                        String Kind = (String) tableModel.getValueAt(row, 4);
                        String Cost = (String) tableModel.getValueAt(row, 5);
                        String Number_books = (String) tableModel.getValueAt(row, 6);
                        pst.setString(8, temp);
                        pst.setString(1, idBook);
                        pst.setString(2, Book_name);
                        pst.setString(3, Author);
                        pst.setString(4, Publisher);
                        pst.setString(5, Kind);
                        pst.setString(6, Cost);
                        pst.setString(7, Number_books);
                        pst.addBatch();
                        pst.executeUpdate();

                        connection.commit();
                    } catch (HeadlessException | SQLException ex) {
                        JOptionPane.showMessageDialog(null, "Can not update!\n" + ex.getMessage());
                    }
                }
                JOptionPane.showMessageDialog(null, "Successfully update");
            }
        });
        insertAbove.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseReleased(MouseEvent e) {
                ((DefaultTableModel) loanPaymentTable.getModel()).insertRow(loanPaymentTable.getSelectedRow(), new Object[]{null});
                loanPaymentTable.removeRowSelectionInterval(loanPaymentTable.getSelectedRow(), loanPaymentTable.getSelectedRow());
            }
        });
        insertBelow.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseReleased(MouseEvent e) {
                ((DefaultTableModel) loanPaymentTable.getModel()).insertRow(loanPaymentTable.getSelectedRow() + 1, new Object[]{null});
                loanPaymentTable.removeRowSelectionInterval(loanPaymentTable.getSelectedRow() + 1, loanPaymentTable.getSelectedRow() + 1);
            }
        });
        subPopupMenu.add(deleteFromTb);
        subPopupMenu.add(deleteFromDb);
        insertMenu.add(insertBelow);
        insertMenu.add(insertAbove);
        deleteMenu.add(deleteFromTb);
        deleteMenu.add(deleteFromDb);
        popupMenu.add(insertMenu);
        popupMenu.add(deleteMenu);
        popupMenu.add(update);
        return popupMenu;

    }

    private boolean isEmptyRow(int row, JTable table) {
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        for (int i = 0; i < table.getColumnCount(); i++) {
            String data = (String) table.getValueAt(row, i);
            if (data == null) {
                return true;
            }
        }
        return false;
    }


    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.JCheckBox checkDetailTable;
    private javax.swing.JCheckBox checkLoanTabel;
    private javax.swing.JButton chooseFile;
    private javax.swing.JButton clearAll;
    private javax.swing.JTable detailTable;
    private javax.swing.JButton insertData;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JToolBar jToolBar1;
    private javax.swing.JToolBar jToolBar2;
    private javax.swing.JTable loanPaymentTable;
    private javax.swing.JButton save;
    private javax.swing.JButton showData;
    // End of variables declaration//GEN-END:variables
}
