package InputSheet.Generator;

import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JSeparator;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.border.EmptyBorder;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.application.Platform;
import javafx.embed.swing.JFXPanel;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

@SuppressWarnings("serial")
public class GUI extends JFrame {
	public static File file;
	public static String ScenarioName;
	public static JComboBox<String> comboBoxLblOutPutType;

	//export file dialog
	public static void exportInputSheet() throws IOException {
		new JFXPanel();
		Platform.runLater(()->{
			//File selectedFolder = showExportDialog("Select destination folder");
			//if(selectedFolder != null)
			try {
				if(comboBoxLblOutPutType.getSelectedItem().equals("Manual Automation")) {
					//	XSSFWorkbook[] wbFiles = InputDataFromUTC.generateInputSheet();
					//	File selectedFolder = showExportDialog("Select destination folder");
					//	if(selectedFolder != null) exportSheetFile(selectedFolder,wbFiles);
					//	System.out.println("File exported sucessully to "+selectedFolder.getAbsolutePath());
				}
				else {
					XSSFWorkbook EDIwb = EDI_InputSheet.generateEdiInputSheet();
					File selectedFolder = showExportDialog("Select destination folder");
					if(selectedFolder != null) exportEDISheetFile(selectedFolder, EDIwb);
					System.out.println("File exported sucessully to "+selectedFolder.getAbsolutePath());
				}
			}
			catch(IOException e) {
				System.out.println("failed to export file");
				e.printStackTrace();
			}
		});
	}
	//method to display native directory chooser dialog
	private static File showExportDialog(String title) {
		DirectoryChooser directoryChooser = new DirectoryChooser();
		directoryChooser.setTitle(title);
		return directoryChooser.showDialog(null);
	}
	private static void exportSheetFile(File destinationFolder, XSSFWorkbook[] wbFiles) throws IOException {
		for(XSSFWorkbook wb:wbFiles) {
			File outputFile;
			if(wb == wbFiles[0]) outputFile = new File(destinationFolder, "Claim Data Institutional.xlsx");
			else outputFile = new File(destinationFolder, "Claim Data Professional.xlsx");
			try(FileOutputStream fos = new FileOutputStream(outputFile) ){
				wb.write(fos);
				fos.close();
			}
		}
	}
	private static void exportEDISheetFile(File destinationFolder,XSSFWorkbook EDIwb) throws IOException {
		File outputFile = new File(destinationFolder, "EDI_InputSheet.x1sx");
		try(FileOutputStream fos = new FileOutputStream(outputFile)){
			EDIwb.write(fos);
			fos.close();
		}
	}

	public static JPanelGradient createMainFrame() {

		JPanelGradient contentPane = new JPanelGradient(new Color(254, 238, 206), new Color(204, 236, 255));
		contentPane.setBackground(new Color(255, 245, 238));
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setBounds(100, 100, 525, 550);
		contentPane.setLayout(null);

		JLabel lblHeader = new JLabel("Input Sheet Generator");
		lblHeader.setForeground(new Color(30, 144, 255));
		lblHeader.setFont(new Font("Times New Roman", Font.BOLD, 24));
		lblHeader.setVerticalAlignment(SwingConstants.CENTER);
		lblHeader.setHorizontalAlignment(SwingConstants.CENTER);
		lblHeader.setBounds(100, 13, 294, 24);
		contentPane.add(lblHeader);

		JLabel lblSourceSheet = new JLabel("Source Sheet : ");
		lblSourceSheet.setFont(new Font("Times New Roman", Font.BOLD, 14));
		lblSourceSheet.setBounds(60, 90, 113, 16);
		contentPane.add(lblSourceSheet);

		JComboBox<String> comboBoxSourceSheet = new JComboBox<String>();
		comboBoxSourceSheet.setFont(new Font ("Times New Roman", Font.BOLD, 14));
		comboBoxSourceSheet.setBackground(Color.WHITE);
		comboBoxSourceSheet.setBounds(235, 90, 200, 20);
		comboBoxSourceSheet.addItem("Select Source Sheet");
		comboBoxSourceSheet.addItem("UTC Sheet");
		comboBoxSourceSheet.addItem("Claim Result Sheet");
		contentPane.add(comboBoxSourceSheet);

		JLabel lblOutPutType = new JLabel("OutPut Type : ");
		lblOutPutType.setFont(new Font("Times New Roman", Font.BOLD, 14) );
		lblOutPutType.setBounds(60, 130, 113, 24);
		contentPane.add(lblOutPutType);

		comboBoxLblOutPutType = new JComboBox<String>();
		comboBoxLblOutPutType.setFont(new Font("Times New Roman", Font.BOLD, 14));
		comboBoxLblOutPutType.setBackground(Color.WHITE);
		comboBoxLblOutPutType.setBounds(235, 130, 200, 20);
		comboBoxLblOutPutType.addItem("Select OutPut Type");
		comboBoxLblOutPutType.addItem("Manual Automation");
		comboBoxLblOutPutType.addItem("EDI Automation");
		contentPane.add(comboBoxLblOutPutType);

		JLabel lblScenarioName = new JLabel("Scenario Name : ");
		lblScenarioName.setFont(new Font("Times New Roman", Font.BOLD, 14));
		lblScenarioName.setBounds(60, 170, 113, 32);
		contentPane.add(lblScenarioName);

		JTextField textField_ScenarioName = new JTextField();
		textField_ScenarioName.setBounds(235, 170, 200, 20);
		contentPane.add(textField_ScenarioName);
		textField_ScenarioName.setColumns(10);

		JButton btnImportSheet = new JButton("Import Sheet");
		btnImportSheet.setFont(new Font("Times New Roman", Font.BOLD, 12));
		btnImportSheet.setBounds(60, 210, 113, 20);
		contentPane.add(btnImportSheet);
		btnImportSheet.addActionListener(
				e->{
					new JFXPanel();
					Platform.runLater(()->{
						FileChooser fileChooser = new FileChooser();
						fileChooser.setTitle("Select Source Sheet");
						GUI.file = fileChooser.showOpenDialog(null);
						if(file != null) {
							SwingUtilities.invokeLater(()->{
								JOptionPane. showMessageDialog(contentPane, "selected file : "+file.getAbsolutePath());
							});
						}});
				});

		JButton btnNewButton = new JButton("Generate");
		btnNewButton.setBackground(new Color(127, 255, 212));
		btnNewButton.setFont(new Font("Times New Roman", Font.BOLD, 16));
		btnNewButton.setBounds(60, 400,175, 25);
		contentPane.add(btnNewButton);

		JButton btnNewButton2 = new JButton("EXIT");
		btnNewButton2.setBackground(new Color(255, 192, 203));
		btnNewButton2.setFont(new Font("Times New Roman", Font.BOLD, 16));
		btnNewButton2.setBounds(262, 400, 175,25);
		contentPane.add(btnNewButton2);

		JSeparator separator = new JSeparator();
		separator.setForeground(new Color(30, 144, 255));
		separator.setBackground(new Color(30, 144, 255));
		separator.setBounds(100, 40, 298, 50);
		contentPane.add(separator);

		contentPane.setVisible(true);

		// Start Button
		btnNewButton.addActionListener(e->{
			try {
				ScenarioName = textField_ScenarioName.getText();
				exportInputSheet();
			} catch (IOException el) {
				// TODO Auto-generated catch block
				el.printStackTrace();
			}
		});
		// Exit Button
		btnNewButton2.addActionListener(new ActionListener(){
			@Override
			public void actionPerformed(ActionEvent e) {
				System.exit(0);
			}
		});

		return contentPane;
	}

	public static void main(String[] args) {
		JFrame frame = new JFrame ();
		frame.setTitle("InputSheet Generatort");
		frame.setBounds(100, 100, 525, 550);
		frame.add(createMainFrame());
		frame.setVisible(true);
	}
}

