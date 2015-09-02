/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.schneider.tsmteam;

import java.awt.Color;
import java.awt.GraphicsConfiguration;
import java.io.*;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.awt.FileDialog;
import java.awt.Frame;
import java.io.*;
import javax.swing.*;
import javax.swing.JProgressBar;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



/**
 *
 * @author Carlos Gonzalez
 */
public class MainTSMWindow extends javax.swing.JFrame {
    String textt="";
    String[] listadearchivos = new String[50];
    int contador=0;
   
    public MainTSMWindow() {
        initComponents();
    }

  
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jTextFieldDirectory = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTextAreaFilesList = new javax.swing.JTextArea();
        jButton2 = new javax.swing.JButton();
        jLabelState = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("TSM Team - EMS");
        setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        setMaximumSize(new java.awt.Dimension(400, 500));
        setMinimumSize(new java.awt.Dimension(400, 500));
        setPreferredSize(new java.awt.Dimension(400, 500));
        setResizable(false);

        jPanel1.setBackground(new java.awt.Color(51, 51, 51));

        jButton1.setBackground(new java.awt.Color(0, 153, 51));
        jButton1.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jButton1.setForeground(new java.awt.Color(255, 255, 255));
        jButton1.setText("Search...");
        jButton1.setBorderPainted(false);
        jButton1.setFocusPainted(false);
        jButton1.setMargin(new java.awt.Insets(0, 0, 0, 0));
        jButton1.setMaximumSize(new java.awt.Dimension(58, 23));
        jButton1.setMinimumSize(new java.awt.Dimension(58, 23));
        jButton1.setOpaque(false);
        jButton1.setPreferredSize(new java.awt.Dimension(58, 23));
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jTextFieldDirectory.setEditable(false);
        jTextFieldDirectory.setBackground(new java.awt.Color(204, 204, 204));
        jTextFieldDirectory.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextFieldDirectory.setForeground(new java.awt.Color(102, 102, 102));
        jTextFieldDirectory.setText("Directory...");

        jLabel2.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel2.setText("Technology Standardization Management");
        jLabel2.setToolTipText("");

        jLabel3.setFont(new java.awt.Font("Tahoma", 0, 24)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel3.setText("EMS");
        jLabel3.setToolTipText("");

        jTextAreaFilesList.setEditable(false);
        jTextAreaFilesList.setBackground(new java.awt.Color(204, 204, 204));
        jTextAreaFilesList.setColumns(20);
        jTextAreaFilesList.setForeground(new java.awt.Color(102, 102, 102));
        jTextAreaFilesList.setRows(5);
        jScrollPane1.setViewportView(jTextAreaFilesList);

        jButton2.setBackground(new java.awt.Color(0, 153, 51));
        jButton2.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jButton2.setForeground(new java.awt.Color(255, 255, 255));
        jButton2.setText("Start!");
        jButton2.setBorderPainted(false);
        jButton2.setFocusPainted(false);
        jButton2.setMargin(new java.awt.Insets(0, 0, 0, 0));
        jButton2.setMaximumSize(new java.awt.Dimension(58, 23));
        jButton2.setMinimumSize(new java.awt.Dimension(58, 23));
        jButton2.setOpaque(false);
        jButton2.setPreferredSize(new java.awt.Dimension(58, 23));
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jLabelState.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jLabelState.setForeground(new java.awt.Color(153, 153, 153));

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(0, 0, Short.MAX_VALUE)
                                .addComponent(jLabel2)))
                        .addGap(29, 29, 29))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addComponent(jLabelState, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 91, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addGroup(jPanel1Layout.createSequentialGroup()
                                    .addComponent(jTextFieldDirectory, javax.swing.GroupLayout.PREFERRED_SIZE, 282, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 91, javax.swing.GroupLayout.PREFERRED_SIZE))
                                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 379, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addContainerGap(18, Short.MAX_VALUE))))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel2)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTextFieldDirectory, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 39, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 305, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabelState, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(62, 62, 62))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
     
        jTextAreaFilesList.setText("");                                                             // Limpia la lista de Archivos en el jTextArea
        jTextFieldDirectory.setText("");
        JFileChooser jFileChooserFilesDirectory = new JFileChooser();                               // Se crea el JFileChooser
        jFileChooserFilesDirectory.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);             // Establece que el JFileChooser solo leera directorios
        int seleccion=jFileChooserFilesDirectory.showOpenDialog(jLabel2);
        if(seleccion==JFileChooser.APPROVE_OPTION)
            {
                File archivo = jFileChooserFilesDirectory.getSelectedFile();
                jTextFieldDirectory.setText(archivo.getAbsolutePath());
            } 
         
        ObtenerListado(jTextFieldDirectory.getText());  // Se llama al metodo obtener listado
        
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
      ProcesaXLS();        // TODO add your handling code here:
      
      
      
     
      jButton2.setVisible(false);
      jButton1.setVisible(false);
    }//GEN-LAST:event_jButton2ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainTSMWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainTSMWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainTSMWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainTSMWindow.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainTSMWindow().setVisible(true);
             
            }
        });
    }

   
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabelState;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextArea jTextAreaFilesList;
    private javax.swing.JTextField jTextFieldDirectory;
    // End of variables declaration//GEN-END:variables

  
    
    private void ObtenerListado(String text)    {

        File folder;
        folder = new File(text);
        File[] listOfFiles = folder.listFiles();

        for (File listOfFile : listOfFiles) 
        {
            if (listOfFile.isFile()) 
            {
               System.out.println("File " + listOfFile.getName());
                textt=jTextAreaFilesList.getText();
                if(listOfFile.getName().endsWith("all.xls")||listOfFile.getName().endsWith("upd.xls") ){
                jTextAreaFilesList.setText(textt + listOfFile.getName() + "\n");
                listadearchivos[contador]=listOfFile.getName();
                contador++;}
            } 
        }
    }
    
    
    
    private void ProcesaXLS() {
        
        final SwingWorker worker = new SwingWorker() {
            
            @Override
            protected Object doInBackground() throws Exception {
                String contenido = "s";
                int unidadProgresBAR=contador/100;
                for (int   uy= 0; uy < contador; uy++) {
                    jLabelState.setText(listadearchivos[uy]+" is in Processing");
                    
                    
                    
                     if(listadearchivos[uy].endsWith("all.xls")){
        System.out.println(listadearchivos[uy] + "Termina con All");
        
        try {
                       
                        FileInputStream file = new FileInputStream(new File(jTextFieldDirectory.getText()+ "\\" + listadearchivos[uy]));
                        HSSFWorkbook workbook = new HSSFWorkbook(file);
                        HSSFSheet sheet = workbook.getSheetAt(0);
                        Cell cell = null;
                        int sheetsize = sheet.getPhysicalNumberOfRows() ;
                        
                        
                        cell = sheet.getRow(0).getCell(21);
                        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                        if (sheetsize>1){
                           
                            
                                
                            
                        cell = sheet.getRow(1).getCell(21);
                      
                        for (int i=1;i<sheetsize;i++){
                            cell = sheet.getRow(i).getCell(21);
                            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
                            {
                                String cellContents = cell.getStringCellValue();
                                cell = sheet.getRow(i).getCell(1);
                                cell.setCellValue(cellContents);
                                cell = sheet.getRow(i).getCell(21);
                                cell.setCellType(Cell.CELL_TYPE_BLANK);
                            }
                        }
                        }
                         
                        cell = sheet.getRow(0).getCell(21);
                        cell.setCellType(Cell.CELL_TYPE_BLANK);
                        }
                        file.close();
                       // ORIGINAL FileOutputStream outFile =new FileOutputStream(new File(jTextFieldDirectory.getText()+ "\\" + listadearchivos[uy]));
                      /*  TEST*/  FileOutputStream outFile =new FileOutputStream(new File(jTextFieldDirectory.getText()+ "\\" + listadearchivos[uy]));
                        workbook.write(outFile);
                        outFile.close();
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                      
             } 
    else {
        System.out.println(listadearchivos[uy] + "Termina con Upd");
        
        try {
                       
                        FileInputStream file = new FileInputStream(new File(jTextFieldDirectory.getText()+ "\\" + listadearchivos[uy]));
                        HSSFWorkbook workbook = new HSSFWorkbook(file);
                        HSSFSheet sheet = workbook.getSheetAt(0);
                        Cell cell = null;
                        int sheetsize = sheet.getPhysicalNumberOfRows() ;
                        cell = sheet.getRow(0).getCell(22);
                        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
                        if (sheetsize>1){
                        cell = sheet.getRow(1).getCell(22);
                        for (int i=1;i<sheetsize;i++){
                            cell = sheet.getRow(i).getCell(22);
                            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
                            {
                                String cellContents = cell.getStringCellValue();
                                cell = sheet.getRow(i).getCell(1);
                                cell.setCellValue(cellContents);
                                cell = sheet.getRow(i).getCell(22);
                                cell.setCellType(Cell.CELL_TYPE_BLANK);
                            }
                        }
                        }
                        cell = sheet.getRow(0).getCell(22);
                        cell.setCellType(Cell.CELL_TYPE_BLANK);
                        }
                        file.close();
                        FileOutputStream outFile =new FileOutputStream(new File(jTextFieldDirectory.getText()+ "\\" + listadearchivos[uy]));
                        workbook.write(outFile);
                        
                        outFile.close();
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
            }

                    jLabelState.setText(listadearchivos[uy]+" is Completed");
                    Thread.sleep(300);
                }
                JOptionPane hola = new JOptionPane();
                jButton2.setVisible(true);
                jButton1.setVisible(true);
                jLabelState.setText("Complete");
                JOptionPane.showMessageDialog(hola,"Complete");
                System.out.println(contador);
                return null;
            }
            
        };
        
        worker.execute();   
    }
}
