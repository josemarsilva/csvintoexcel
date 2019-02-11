package org.csvintoexcel.cli;

import java.io.File;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.commons.lang3.StringUtils;


/*
 * CliArgsParser class is responsible for extract, compile and check consistency
 * for command line arguments passed as parameters in command line
 * 
 * @author josemarsilva
 * @date   2019-02-02
 * 
 */
public class CliArgsParser {
	
	// Constants ...
	public final static String APP_NAME = new String("csvintoexcel");
	public final static String APP_VERSION = new String("v.2019.02.11");
	public final static String APP_USAGE = new String(APP_NAME + " [<args-options-list>] - "+ APP_VERSION);

	// Constants defaults ...
	public final static int DEFAULT_EXCEL_SHEET_NUMBER = 0;
	public final static int DEFAULT_EXCEL_ROW_NUMBER = 1;
	public final static int DEFAULT_EXCEL_COLUMN_NUMBER = 0;
	public final static int DEFAULT_CSV_FILE_IGNORE_HEADER = 1;

	// Constants Error Messages ...
	public final static String MSG_ERROR_FILE_NOT_FOUND = new String("Erro: arquivo '%s' não existe !");
	
	// Properties ...
    private String inputExcelFileOption = new String("");
    private int inputExcelSheetNumberOption = DEFAULT_EXCEL_SHEET_NUMBER;
    private int inputExcelRowNumberOption = DEFAULT_EXCEL_ROW_NUMBER;
    private int inputExcelColumnNumberOption = DEFAULT_EXCEL_COLUMN_NUMBER;
    private String inputExcelDataTypeList = new String("");
    private String inputCsvFileOption = new String("");
    private int inputCsvFileIgnoreHeaderOption = DEFAULT_CSV_FILE_IGNORE_HEADER;
    private String outputExcelFileOption = new String("");


    /*
     * CliArgsParser(args) - Constructor ...
     */
	public CliArgsParser( String[] args ) {

		// Options creating ...
		Options options = new Options();
		
		
		// Options configuring  ...
        Option helpOption = Option.builder("h") 
        		.longOpt("help") 
        		.required(false) 
        		.desc("shows usage help message. See more https://github.com/josemarsilva/csvintoexcel") 
        		.build(); 
        Option inputExcelFileOption = Option.builder("e")
        		.longOpt("input-excel-file") 
        		.required(true) 
        		.desc("Nome do arquivo que contem Pasta de trabalho EXCEL (.xls ou .xlsx) usada na juncao. Ex: template.xlsx")
        		.hasArg()
        		.build();
        Option inputExcelSheetNumberOption = Option.builder("s")
                .longOpt("input-excel-sheet-number") 
                .required(false) 
                .type(Number.class)
                .desc("Numero sequencial da PLANILHA dentro da Pasta de trabalho usada na juncao. Ex: 0-n (0=primeira). Default 0") 
        		.hasArg()
                .build(); 
        Option inputExcelRowNumberOption = Option.builder("r")
                .longOpt("input-excel-row-number") 
                .required(false) 
                .type(Number.class)
                .desc("Numero da LINHA inicial da planilha usada na juncao. Ex: 0-n (0=primeira). Default 2") 
        		.hasArg()
                .build(); 
        Option inputExcelColumnNumberOption = Option.builder("c")
                .longOpt("input-excel-column-number") 
                .required(false) 
                .type(Number.class)
                .desc("Numero da COLUNA inicial da planilha. Ex: 0-n (0=primeira). Default 0") 
        		.hasArg()
                .build(); 
        Option inputExcelDataTypeList = Option.builder("d")
        		.longOpt("input-excel-data-type-list") 
        		.required(false) 
        		.desc("Lista dos tipos de dados (data-type) das células separados por '-' conforme 'https://dzone.com/articles/java-string-format-examples'. Ex: %s-%d-%f")
        		.hasArg()
        		.build();
        Option inputCsvFileOption = Option.builder("f")
                .longOpt("input-csv-file") 
                .required(true) 
                .desc("Nome do arquivo (.csv) que contem o conteúdo usado na juncao. Ex: dados.csv (separador ;)") 
        		.hasArg()
                .build(); 
        Option inputCsvFileIgnoreHeaderOption = Option.builder("g")
                .longOpt("input-csv-file-ignore-header") 
                .required(false) 
                .type(Number.class)
                .desc("Numero de LINHAS DE CABECALHO dos dados ignorados na juncao. Ex: 1 (Default e 1)") 
        		.hasArg()
                .build(); 
        Option outputExcelFileOption = Option.builder("o")
                .longOpt("output-excel-file") 
                .required(true) 
        		.desc("Nome do arquivo que da Pasta de trabalho EXCEL (.xls ou .xlsx) conter a juncao. Ex: template-com-dados.xlsx")
        		.hasArg()
                .build(); 
        
        
		// Options adding configuration ...
        options.addOption(helpOption);
        options.addOption(inputExcelFileOption);
        options.addOption(inputExcelSheetNumberOption);
        options.addOption(inputExcelRowNumberOption);
        options.addOption(inputExcelColumnNumberOption);
        options.addOption(inputExcelDataTypeList);
        options.addOption(inputCsvFileOption);
        options.addOption(inputCsvFileIgnoreHeaderOption);
        options.addOption(outputExcelFileOption);
        
        
        // CommandLineParser ...
        CommandLineParser parser = new DefaultParser();
		try {
			CommandLine cmdLine = parser.parse(options, args);
			
	        if (cmdLine.hasOption("help")) { 
	            HelpFormatter formatter = new HelpFormatter();
	            formatter.printHelp(APP_USAGE, options);
	            System.exit(0);
	        } else {
	        	
	        	// Set properties from Options ...
	        	this.setInputExcelFileOption( cmdLine.getOptionValue("input-excel-file", "") );	        	
	        	this.setInputExcelSheetNumberOption( (cmdLine.getParsedOptionValue("input-excel-sheet-number")==null) ? DEFAULT_EXCEL_SHEET_NUMBER : Integer.parseInt( cmdLine.getParsedOptionValue("input-excel-sheet-number").toString() ) );
	        	this.setInputExcelRowNumberOption( (cmdLine.getParsedOptionValue("input-excel-row-number")==null) ? DEFAULT_EXCEL_ROW_NUMBER : Integer.parseInt( cmdLine.getParsedOptionValue("input-excel-row-number").toString() ) );
	        	this.setInputExcelColumnNumberOption( (cmdLine.getParsedOptionValue("input-excel-column-number")==null) ? DEFAULT_EXCEL_COLUMN_NUMBER : Integer.parseInt( cmdLine.getParsedOptionValue("input-excel-column-number").toString() ) );
	        	this.setInputExcelDataTypeList( cmdLine.getOptionValue("input-excel-data-type-list", "") );	        	
	        	this.setInputCsvFileOption( cmdLine.getOptionValue("input-csv-file", "") );
	        	this.setInputCsvFileIgnoreHeaderOption( (cmdLine.getParsedOptionValue("input-csv-file-ignore-header")==null) ? DEFAULT_CSV_FILE_IGNORE_HEADER : Integer.parseInt( cmdLine.getParsedOptionValue("input-csv-file-ignore-header").toString() ) );
	        	this.setOutputExcelFileOption( cmdLine.getOptionValue("output-excel-file", "") );
	        	
	        	// Check arguments Options ...
	        	try {
	        		checkArgumentOptions();
	        	} catch (Exception e) {
	    			System.err.println(e.getMessage());
	    			System.exit(-1);
	        	}
	        	
	        }
			
		} catch (ParseException e) {
			System.err.println(e.getMessage());
            HelpFormatter formatter = new HelpFormatter(); 
            formatter.printHelp(APP_USAGE, options); 
			System.exit(-1);
		} 
        
	}

	//
	private void checkArgumentOptions() throws Exception {
		
		// Check input file arguments must exists ...
		if (!(new File(inputExcelFileOption).exists())) {
			throw new Exception(MSG_ERROR_FILE_NOT_FOUND.replaceFirst("%s", inputExcelFileOption));
		} else if (!(new File(inputCsvFileOption).exists())) {
			throw new Exception(MSG_ERROR_FILE_NOT_FOUND.replaceFirst("%s", inputCsvFileOption));
		}
		
	}
	
	// Getters and Setters ...

	public String getInputExcelFileOption() {
		return inputExcelFileOption;
	}


	public void setInputExcelFileOption(String inputExcelFileOption) {
		this.inputExcelFileOption = inputExcelFileOption;
	}


	public int getInputExcelSheetNumberOption() {
		return inputExcelSheetNumberOption;
	}


	public void setInputExcelSheetNumberOption(int inputExcelSheetNumberOption) {
		this.inputExcelSheetNumberOption = inputExcelSheetNumberOption;
	}


	public int getInputExcelRowNumberOption() {
		return inputExcelRowNumberOption;
	}


	public void setInputExcelRowNumberOption(int inputExcelRowNumberOption) {
		this.inputExcelRowNumberOption = inputExcelRowNumberOption;
	}


	public int getInputExcelColumnNumberOption() {
		return inputExcelColumnNumberOption;
	}


	public void setInputExcelColumnNumberOption(int inputExcelColumnNumberOption) {
		this.inputExcelColumnNumberOption = inputExcelColumnNumberOption;
	}


	public String getInputExcelDataTypeList() {
		return this.inputExcelDataTypeList;
	}

	
	private void setInputExcelDataTypeList(String optionValue) {
		this.inputExcelDataTypeList = optionValue;
	}

	
	public String getInputCsvFileOption() {
		return inputCsvFileOption;
	}


	public void setInputCsvFileOption(String inputCsvFileOption) {
		this.inputCsvFileOption = inputCsvFileOption;
	}


	public int getInputCsvFileIgnoreHeaderOption() {
		return inputCsvFileIgnoreHeaderOption;
	}


	public void setInputCsvFileIgnoreHeaderOption(int inputCsvFileIgnoreHeaderOption) {
		this.inputCsvFileIgnoreHeaderOption = inputCsvFileIgnoreHeaderOption;
	}


	public String getOutputExcelFileOption() {
		return outputExcelFileOption;
	}


	public void setOutputExcelFileOption(String outputExcelFileOption) {
		this.outputExcelFileOption = outputExcelFileOption;
	}
	


}
