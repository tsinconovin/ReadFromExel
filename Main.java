import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        String[][] array = new String[10][7];
        Integer count_r = 0;
        Integer count_c = 0;

        try {
            File file = new File("C:\\Users\\tsinco-pc10\\Documents\\MyProject\\Source File\\Excel-3692.xlsx");   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            //creating Workbook instance that refers to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            for (int i = 0; i <= 9; i++) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                for (int j = 0; j <= 5; j++) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                            System.out.print(cell.getStringCellValue() + "\t\t\t");
                            array[i][j] = cell.getStringCellValue();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type
                            System.out.print(cell.getNumericCellValue() + "\t\t\t");
                            array[i][j] = cell.getStringCellValue();
                            break;
                        default:
                    }
                    count_c++;
                }
                System.out.println("");
                count_r++;
            }


            JFrame frame = new JFrame();

            String[] columnNames = {"کدملی", "نام", "نام خانوادگی", "نام پدر", "ناریخ تولد", "شماره شناسنامه"};

            JTable table = new JTable(array, columnNames);

            DefaultTableModel model;
            model = new DefaultTableModel(array, columnNames);
            JButton add = new JButton("Add");
            table = new JTable();
            table.setModel(model);
            add.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent ae) {
                    model.addRow(array[0]);
                }
            });
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            Container c = frame
                    .getContentPane();
            c.setLayout(new BorderLayout());
            c.add(add, BorderLayout.SOUTH);
            c.add(new JScrollPane(table), BorderLayout.CENTER);

            frame.pack();

            frame.setLocationRelativeTo(null);



            frame.add(table);
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(400, 400);
            Color ivory = new Color(22, 22, 101);
            table.setBackground(ivory);
            table.setForeground(Color.white);
            table.setFont(new Font("Serif", Font.PLAIN, 20));
            frame.setLocationRelativeTo(null);
            table.setRowHeight(50);


            setCellsAlignment(table,SwingConstants.CENTER);



            frame.setVisible(true);


        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    public static void setCellsAlignment(JTable table, int alignment) {
        DefaultTableCellRenderer rightRenderer = new DefaultTableCellRenderer();
        rightRenderer.setHorizontalAlignment(alignment);

        TableModel tableModel = table.getModel();

        for (int columnIndex = 0; columnIndex < tableModel.getColumnCount(); columnIndex++) {
            table.getColumnModel().getColumn(columnIndex).setCellRenderer(rightRenderer);
        }
    }
}
