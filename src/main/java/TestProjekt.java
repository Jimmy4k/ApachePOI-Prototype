import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import resources.CostTableEntry;
import resources.CostType;

public class TestProjekt extends JFrame {
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private JTable table;
	private JTextField tfType;
	private JTextField tfCost;
	private JTextField tfMargin;
	private DefaultTableModel model;
	private CostType ct;
	private CostTableEntry cTE;
	private Map<CostType, CostTableEntry> costMap;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					TestProjekt frame = new TestProjekt();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
	/**
	 * Create the frame.
	 */
	public TestProjekt() {
		setTitle("Export-Import Values to Excel");
		costMap = new HashMap<CostType, CostTableEntry>();
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 758, 447);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		JButton btnExportButton = new JButton("Export to Excel-file");
		btnExportButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				exportToExcel();
			}
		});
		btnExportButton.setBounds(567, 271, 142, 23);
		contentPane.add(btnExportButton);

		new CostType();
		new CostTableEntry();

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(29, 37, 689, 208);
		contentPane.add(scrollPane);

		table = new JTable();
		table.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "Name", "type", "cost", "costWithMargin", "costMargin", "Margin %" }) {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;
			@SuppressWarnings("rawtypes")
			Class[] columnTypes = new Class[] { String.class, String.class, Double.class, Double.class, Double.class,
					Double.class };

			public Class<?> getColumnClass(int columnIndex) {
				return columnTypes[columnIndex];
			}
		});
		table.getColumnModel().getColumn(3).setPreferredWidth(116);
		scrollPane.setViewportView(table);

		tfType = new JTextField();
		tfType.setBounds(29, 272, 125, 20);
		contentPane.add(tfType);
		tfType.setColumns(10);

		tfCost = new JTextField();
		tfCost.setBounds(164, 272, 117, 20);
		contentPane.add(tfCost);
		tfCost.setColumns(10);

		JLabel lblType = new JLabel("Type:");
		lblType.setBounds(29, 256, 46, 14);
		contentPane.add(lblType);

		JLabel lblCost = new JLabel("Cost:");
		lblCost.setBounds(164, 256, 46, 14);
		contentPane.add(lblCost);

		JLabel lblMargin = new JLabel("Margin: %");
		lblMargin.setBounds(294, 256, 86, 14);
		contentPane.add(lblMargin);

		tfMargin = new JTextField();
		tfMargin.setBounds(294, 272, 86, 20);
		contentPane.add(tfMargin);
		tfMargin.setColumns(10);

		JButton btnEinfügen = new JButton("Einf\u00FCgen");
		btnEinfügen.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				insertMapValues(costMap);
				insertValuesToTable(costMap);
			}
		});
		btnEinfügen.setBounds(390, 271, 142, 23);
		contentPane.add(btnEinfügen);

		JButton btnImportButton = new JButton("Import from Excel-file");
		btnImportButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				insertValuesToTable(importFromExcel(costMap));
			}
		});
		// CECKSTYLE:OFF
		btnImportButton.setBounds(567, 305, 142, 23);
		contentPane.add(btnImportButton);

		JButton btnBeendenButton = new JButton("Beenden");
		btnBeendenButton.addActionListener((e) -> System.exit(0));
		
		btnBeendenButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});
		btnBeendenButton.setBounds(567, 374, 142, 23);
		contentPane.add(btnBeendenButton);
	}

	public Map<CostType, CostTableEntry> insertMapValues(Map<CostType, CostTableEntry> costMap) {
		try {
			if (tfType.getText().isEmpty()) {
				throw new NullPointerException("entry Typ!");
			}
			if (tfCost.getText().isEmpty()) {
				throw new NullPointerException("entry Cost!");
			}
			if (tfMargin.getText().isEmpty()) {
				throw new NullPointerException("entry Margin!");
			} else {
				ct = new CostType();
				cTE = new CostTableEntry();

				if (tfType.getText().matches(".*[a-z].*")) {
					ct.setCostType(tfType.getText());

				}
				if (!tfCost.getText().matches(".*[a-z].*")) {
					cTE.setCost(Double.parseDouble(tfCost.getText()));
				} else {
					throw new NumberFormatException("CostType eingeben!");
				}

				cTE.setMargin(Double.parseDouble(tfMargin.getText()));
				cTE.setCostMargin((cTE.getCost() / 100) * cTE.getMargin());
				cTE.setCostWithMargin(cTE.getCost() + cTE.getCostMargin());
				costMap.put(ct, cTE);
				System.out.println(costMap.size());
			}

		} catch (NullPointerException e) {
			if (tfType.getText().isEmpty()) {
				tfType.setText(e.getMessage());
			}
			if (tfCost.getText().isEmpty()) {
				tfCost.setText(e.getMessage());
			}
			if (tfMargin.getText().isEmpty()) {
				tfMargin.setText(e.getMessage());
			}
		}
		return costMap;
	}

	public void insertValuesToTable(Map<CostType, CostTableEntry> map) {
		model = (DefaultTableModel) table.getModel();
		model.setRowCount(0);
		for (Iterator<Entry<CostType, CostTableEntry>> it = map.entrySet().iterator(); it.hasNext();) {
			Entry<CostType, CostTableEntry> entry = it.next();

			model.addRow(new Object[] { entry.getKey().getCostType(), entry.getKey().getCostType(),
					entry.getValue().getCost(), entry.getValue().getCostWithMargin(), entry.getValue().getCostMargin(),
					entry.getValue().getMargin() });
		}
	}

	public void exportToExcel() {
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("CostList");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell;

		int rowNum;
		int colNum;
		int tempRows;
		int cellnum;
		int rowCount = model.getRowCount();
		int columnCount = model.getColumnCount();

		for (colNum = 0; colNum < columnCount; colNum++) {
			// create cells
			cell = row.createCell(colNum);
			cell.setCellValue(model.getColumnName(colNum));
		}

		for (rowNum = 0; rowNum < rowCount; rowNum++) {
			tempRows = rowNum + 1;
			row = sheet.createRow(tempRows);
			for (cellnum = 0; cellnum < columnCount; cellnum++) {
				cell = row.createCell(cellnum);

				Object value = model.getValueAt(rowNum, cellnum);
				Double val = null;

				if (value instanceof Number) {
					val = ((Number) value).doubleValue();
					cell.setCellValue(val);
				} else {
					cell.setCellValue(model.getValueAt(rowNum, cellnum).toString());
				}
			}
		}

		for (int columnIndex = 0; columnIndex < 10; columnIndex++) {
			sheet.autoSizeColumn(columnIndex);
		}
		try {
			FileOutputStream outputStream = new FileOutputStream("Results.xlsx");
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		System.out.println("Excel file was successful!");
	}

	@SuppressWarnings("incomplete-switch")
	public Map<CostType, CostTableEntry> importFromExcel(Map<CostType, CostTableEntry> costMap) {
		String excelFilePath = ".\\Results.xlsx";
		try {

			costMap.clear();
			FileInputStream inputStream = new FileInputStream(excelFilePath);
			@SuppressWarnings("resource")
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Using for loop
			int rows = sheet.getLastRowNum();
			int cols = sheet.getRow(1).getLastCellNum();

			for (int r = 1; r <= rows; r++) {
				ct = new CostType();
				cTE = new CostTableEntry();
				XSSFRow row = sheet.getRow(r);
				for (int c = 0; c < cols; c++) {
					XSSFCell cell = row.getCell(c);

					switch (cell.getCellType()) {
					case STRING:
						ct.setCostType(cell.getStringCellValue());
						break;
					case NUMERIC:
						if (c == 2) {
							cTE.setCost(cell.getNumericCellValue());
						} else if (c == 5) {
							cTE.setMargin(cell.getNumericCellValue());
							cTE.setCostMargin((cTE.getCost() / 100) * cTE.getMargin());
							cTE.setCostWithMargin(cTE.getCost() + cTE.getCostMargin());
						}
					}
				}
				costMap.put(ct, cTE);
				System.out.println(costMap.size());
			}
		} catch (Exception e) {

		}
		return costMap;
	}
}
