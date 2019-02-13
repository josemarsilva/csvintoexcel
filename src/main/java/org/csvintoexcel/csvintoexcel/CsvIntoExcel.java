package org.csvintoexcel.csvintoexcel;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.csvintoexcel.cli.CliArgsParser;


public class CsvIntoExcel {
	
	// Constants Error Messages ...
	public final static String MSG_ERROR_FILE_COPY = new String("Erro: exception '%s' durante copia de arquivos de '%s' para '%s'!");
	public final static String MSG_ERROR_FILE_SAVE = new String("Erro: exception '%s' durante escrita de arquivos de '%s''!");
	public final static String MSG_ERROR_EXCEL_EXCEPTION = new String("Erro: exception '%s' durante operação '%s' com parametros '%s'!");

	
	// Properties ...
	private CliArgsParser cliArgsParser;
	
	private Workbook workbook;
	private Sheet worksheet;
	private HashMap<Integer, CellStyle> cellStyles = new HashMap<Integer, CellStyle>();


    /*
     * CsvIntoExcel(cliArgsParser) - Constructor ...
     */
	public CsvIntoExcel(CliArgsParser pCliArgsParser) {
		super();
		this.cliArgsParser = pCliArgsParser;
	}
	
	
	/*
	 * cloneInputIntoOutput() - Copy from excel input file to excel output file.
	 */
	public void cloneInputIntoOutput() {
		
		System.out.println("Cloning from '%s' to '%s'".replaceFirst("%s", this.cliArgsParser.getInputExcelFileOption()).replaceFirst("%s", this.cliArgsParser.getOutputExcelFileOption()));
		
		Path from = Paths.get(this.cliArgsParser.getInputExcelFileOption());
		Path to = Paths.get(this.cliArgsParser.getOutputExcelFileOption());
		try {
			Files.copy(from, to, StandardCopyOption.REPLACE_EXISTING);
		} catch (IOException e) {
			//e.printStackTrace();
			System.err.println(MSG_ERROR_FILE_COPY.replaceFirst("%s", e.getMessage()).replaceFirst("%s", this.cliArgsParser.getInputExcelFileOption()).replaceFirst("%s", this.cliArgsParser.getOutputExcelFileOption()) );
			System.exit(-1);
		}
		
	}
	
	
	/*
	 * copyFormatStyle() - Copy Formats / Styles from first row.
	 */
	public void copyFormatStyle() {
		String msgCopy = new String("Copy formats/styles");
		System.out.println(msgCopy + " ...");
		
		// Open Workbook ...
		System.out.print (msgCopy+" - opening workbook '%s' ...".replaceFirst("%s",this.cliArgsParser.getOutputExcelFileOption()));
		try {
			this.workbook = WorkbookFactory.create(new FileInputStream(this.cliArgsParser.getOutputExcelFileOption().replace("\\", "//")));
		} catch (Exception e) {
			// e.printStackTrace();
			System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "WorkbookFactory.create()").replaceFirst("%s", this.cliArgsParser.getOutputExcelFileOption().replace("\\", "//") ));
			System.exit(-1);
		}
		
		// Open Worksheet ..
		System.out.println (""+msgCopy+" - opening worksheet index-position (%s) ...".replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelSheetNumberOption()) ));
		this.worksheet = this.workbook.getSheetAt(this.cliArgsParser.getInputExcelSheetNumberOption());
		if (this.worksheet==null) {
			try {
				throw new Exception("worksheet IS NULL");
			} catch (Exception e) {
				// e.printStackTrace();
				System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "workbook.getSheetAt()").replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelSheetNumberOption()) ));
				System.exit(-1);
			}
		}
		
		// Getting first row ..
		System.out.print (""+msgCopy+" - row[%s] ... ".replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelRowNumberOption())));
		Row row = worksheet.getRow( this.cliArgsParser.getInputExcelRowNumberOption() );
		if (row==null) {
			try {
				throw new Exception("row IS NULL");
			} catch (Exception e) {
				// e.printStackTrace();
				System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "worksheet.getRow()").replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelRowNumberOption()) ));
				System.exit(-1);
			}
		}
		
		// Getting cell ..
		for (int j=this.cliArgsParser.getInputExcelColumnNumberOption();j<=row.getLastCellNum();j++) {
			System.out.print ("cell[%s,%s], ".replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelRowNumberOption())).replaceFirst("%s", String.valueOf(j) ));
			Cell cell = row.getCell(j);
			if (cell!=null) {
				CellStyle cellStyle = cell.getCellStyle();
				if (cellStyle!=null) {
					cellStyles.put(j,cellStyle);
				}
			}
		}
		
		System.out.println("");
	}
	
	
	/*
	 * cleanContent() - Clear all contents from firstrow and remove all remaining rows.
	 */
	public void cleanContent() {
		String msgCopy = new String("Clean contents");
		System.out.println(msgCopy + " ...");
		
		// Getting first row ..
		System.out.print (""+msgCopy+" - row[%s] ... ".replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelRowNumberOption())));
		Row row = worksheet.getRow( this.cliArgsParser.getInputExcelRowNumberOption() );
		if (row==null) {
			try {
				throw new Exception("row IS NULL");
			} catch (Exception e) {
				// e.printStackTrace();
				System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "worksheet.getRow()").replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelRowNumberOption()) ));
				System.exit(-1);
			}
		}

		
		// Cleaning all cells from first rows ...
		for (int j=this.cliArgsParser.getInputExcelColumnNumberOption();j<=row.getLastCellNum();j++) {
			System.out.print ("cell[%s,%s], ".replaceFirst("%s", String.valueOf(this.cliArgsParser.getInputExcelRowNumberOption())).replaceFirst("%s", String.valueOf(j) ));
			Cell cell = row.getCell(j);
			if (cell!=null) {
				cell.setCellValue("");
				cell.setCellType(Cell.CELL_TYPE_BLANK);
			}
		}
		System.out.println("");
		
		
		// Remove all remaining rows after first row ...
		for (int i=this.worksheet.getLastRowNum();i>this.cliArgsParser.getInputExcelRowNumberOption();i--) {
			System.out.println (""+msgCopy+" - row[%s] ...".replaceFirst("%s", String.valueOf(i) ));
			Row rowToRemove = worksheet.getRow( i );
			worksheet.removeRow(rowToRemove);
		}
	}
	
	
	/*
	 * copyContents() - Copy contents from input to output.
	 */
	public void copyContents() {
		String msgCopy = new String("Copy contents");
		System.out.println(msgCopy + " ...");
		// Getting Array List of datatype ...
		ArrayList aDataTypeList = null;
		if (cliArgsParser.getInputExcelDataTypeList() != null ) {
			aDataTypeList = new ArrayList(Arrays.asList(cliArgsParser.getInputExcelDataTypeList().split("-")));
		}
		// Reading file ...
		BufferedReader br = null;
		FileReader fr = null;
		int currentLineNumber = 0;
		int nRowNumWorksheet = this.cliArgsParser.getInputExcelRowNumberOption();
		try {
			fr = new FileReader(this.cliArgsParser.getInputCsvFileOption());
			br = new BufferedReader(fr);
			String sCurrentLine;
			// Loop content rows ...
			while ((sCurrentLine = br.readLine()) != null) {
				if (currentLineNumber >= this.cliArgsParser.getInputCsvFileIgnoreHeaderOption()) {
					if ( 
							(currentLineNumber < 10)
							|| (currentLineNumber >= 10 && currentLineNumber < 100 && currentLineNumber % 10 == 0)
							|| (currentLineNumber >= 100 && currentLineNumber < 1000 && currentLineNumber % 100 == 0)
							|| (currentLineNumber >= 1000 && currentLineNumber < 10000 && currentLineNumber % 1000 == 0)
							|| (currentLineNumber >= 10000 && currentLineNumber < 100000 && currentLineNumber % 5000 == 0)
							|| (currentLineNumber >= 100000 && currentLineNumber % 10000 == 0)
						) {
						System.out.println(msgCopy + " - '%s'[%s] -> '%s'[%s]".replaceFirst("%s", this.cliArgsParser.getInputCsvFileOption()).replaceFirst("%s", String.valueOf(currentLineNumber)).replaceFirst("%s", this.cliArgsParser.getOutputExcelFileOption()).replaceFirst("%s", String.valueOf(nRowNumWorksheet) ));
					}
					// Get a list with each csv value ..
					ArrayList aCsvValues= new ArrayList(Arrays.asList(sCurrentLine.split(";")));
					int nCsvValues = aCsvValues.size();
					// Get row destination ...
					Row row = this.worksheet.getRow(nRowNumWorksheet);
					if (row==null) {
						row = this.worksheet.createRow(nRowNumWorksheet);
						if (row==null) {
							try {
								throw new Exception("create row failed and row IS NULL");
							} catch (Exception e) {
								// e.printStackTrace();
								System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "worksheet.createRow()").replaceFirst("%s", String.valueOf(nRowNumWorksheet) ));
								System.exit(-1);
							}
						}
					}
					// Loop each csv values ...
					for (int j=0;j<nCsvValues;j++) {
						String csvValue = (String) aCsvValues.get(j);
						Cell cell = row.getCell(j + this.cliArgsParser.getInputExcelColumnNumberOption());
						if (cell==null) {
							cell = row.createCell(j);
							if (cell==null) {
								try {
									throw new Exception("create cell failed and cell IS NULL");
								} catch (Exception e) {
									// e.printStackTrace();
									System.err.println(MSG_ERROR_EXCEL_EXCEPTION.replaceFirst("s", e.getMessage()).replaceFirst("%s", "worksheet.createCell()").replaceFirst("%s", ( "[" + String.valueOf(nRowNumWorksheet) + "," + String.valueOf(j) + "]" ) ));
									System.exit(-1);
								}
							}
						}
						// Set Value ...
						cell.setCellValue(csvValue); // as String
						boolean bCanSetFormatStyle = true;
						// Set DataType ...
						if (csvValue!=null) {
							if (!csvValue.equals("")) {
								if(j<aDataTypeList.size()) {
									if (aDataTypeList.get(j) != null) {
										if (aDataTypeList.get(j).equals("%d")) {
											bCanSetFormatStyle = false;
											// %d - int, double, bigint ...
											try {
												cell.setCellValue(new BigDecimal(csvValue).doubleValue());
												cell.setCellType(Cell.CELL_TYPE_NUMERIC);
											} catch(Exception e) {
												// Sorry! Could't convert as BigDecimal
											}
										} else if (aDataTypeList.get(j).equals("%f")) {
											bCanSetFormatStyle = false;
											// %f - float, decimal, etc ...
											try {
												cell.setCellValue(new BigDecimal(csvValue.replaceAll(",", ".")).doubleValue());
												cell.setCellType(Cell.CELL_TYPE_NUMERIC);
											} catch(Exception e) {
												// Sorry! Could't convert as BigDecimal
											}
										}
									}
								}
							}
						}
						// Set Format / Style ...
						if (bCanSetFormatStyle && j<cellStyles.size()) {
							CellStyle cellStyle = cellStyles.get(j);
							if (cellStyle != null) {
								cell.setCellStyle(cellStyle);
							}
						}
						
					}
					nRowNumWorksheet++;
				}
				currentLineNumber++;
			}
			System.out.println(msgCopy + " - '%s'[%s] -> '%s'[%s]".replaceFirst("%s", this.cliArgsParser.getInputCsvFileOption()).replaceFirst("%s", String.valueOf(currentLineNumber)).replaceFirst("%s", this.cliArgsParser.getOutputExcelFileOption()).replaceFirst("%s", String.valueOf(nRowNumWorksheet) ));
		} catch (IOException e) {
			System.err.println(e.getMessage());
			System.exit(-1);
		} finally {
			try {
				if (br != null)
					br.close();
				if (fr != null)
					fr.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
		
	}
	
	
	/*
	 * saveWorkbook() - Salva o Workbook
	 */
	public void saveWorkbook() {
		try {
			FileOutputStream fileOutputStream = new FileOutputStream(this.cliArgsParser.getOutputExcelFileOption().replace("\\", "//"));
			this.workbook.write(fileOutputStream);
			fileOutputStream.close();
		} catch (Exception e) {
			//e.printStackTrace();
			System.err.println(MSG_ERROR_FILE_SAVE.replaceFirst("%s", e.getMessage()).replaceFirst("%s", this.cliArgsParser.getInputExcelFileOption()) );
			System.exit(-1);
		}
	}


}
