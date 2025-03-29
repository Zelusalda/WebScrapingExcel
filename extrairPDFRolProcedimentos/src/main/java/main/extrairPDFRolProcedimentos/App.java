package main.extrairPDFRolProcedimentos;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import technology.tabula.Table;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;
import technology.tabula.Page;
import technology.tabula.ObjectExtractor;
import technology.tabula.RectangularTextContainer;

public class App {

    public static void main(String[] args) {
        createDownloadDirectory();
        List<String> pdfUrls = new ArrayList<>();
        pdfUrls.add(
                "https://www.gov.br/ans/pt-br/acesso-a-informacao/participacao-da-sociedade/atualizacao-do-rol-de-procedimentos/Anexo_I_Rol_2021RN_465.2021_RN627L.2024.pdf");
        List<String> downloadedFiles = new ArrayList<>();

        for (int i = 0; i < pdfUrls.size(); i++) {
            String saveDir = "downloads/anexo" + (i + 1) + ".pdf";
            downloadFile(pdfUrls.get(i), saveDir);
            downloadedFiles.add(saveDir);
        }

        if (!downloadedFiles.isEmpty()) {
            String pdfFilePath = downloadedFiles.get(0);
            String excelFilePath = "dados.xlsx";
            extrairDadosPDFParaExcel(pdfFilePath, excelFilePath);
            zipFiles(new String[] { excelFilePath }, "Teste_{José_Luiz_Saldanha_Filho}.zip");
        }
    }

    public static void createDownloadDirectory() {
        File directory = new File("downloads");
        if (!directory.exists()) {
            directory.mkdir();
            System.out.println("Diretório 'downloads' criado.");
        }
    }

    public static void downloadFile(String buscaURL, String saveDir) {
        try (InputStream in = new URL(buscaURL).openStream(); FileOutputStream out = new FileOutputStream(saveDir)) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = in.read(buffer)) != -1) {
                out.write(buffer, 0, bytesRead);
            }
            System.out.println("Arquivo baixado: " + saveDir);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void zipFiles(String[] srcFiles, String zipFile) {
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(zipFile))) {
            for (String srcFile : srcFiles) {
                try (FileInputStream fis = new FileInputStream(srcFile)) {
                    ZipEntry zipEntry = new ZipEntry(new File(srcFile).getName());
                    zos.putNextEntry(zipEntry);
                    byte[] buffer = new byte[1024];
                    int length;
                    while ((length = fis.read(buffer)) >= 0) {
                        zos.write(buffer, 0, length);
                    }
                    zos.closeEntry();
                }
            }
            System.out.println("Arquivos compactados em: " + zipFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void extrairDadosPDFParaExcel(String pdfFilePath, String excelFilePath) {
        try (PDDocument document = PDDocument.load(new File(pdfFilePath));
             FileOutputStream fileOut = new FileOutputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("Dados");

            ObjectExtractor extractor = new ObjectExtractor(document);
            SpreadsheetExtractionAlgorithm sea = new SpreadsheetExtractionAlgorithm();


            String[] header = {
                    "PROCEDIMENTO", "RN (alteração)", "VIGÊNCIA", "OD (Seg. Odontológica)", "AMB (Seg. Ambulatorial)", "HCO", "HSO",
                    "REF", "PAC", "DUT", "SUBGRUPO", "GRUPO", "CAPÍTULO"
            };


            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            int rowNum = 0;
            for (int pageNumber = 1; pageNumber <= 181; pageNumber++) {
                Page page = extractor.extract(pageNumber);
                List<Table> tables = sea.extract(page);

                for (Table table : tables) {
                    Row headerRow = sheet.createRow(rowNum++);
                    for (int i = 0; i < header.length; i++) {
                        Cell cell = headerRow.createCell(i);
                        cell.setCellValue(header[i]);
                        cell.setCellStyle(headerStyle);
                    }

                    List<List<RectangularTextContainer>> rows = table.getRows();


                    for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
                        List<RectangularTextContainer> row = rows.get(rowIndex);
                        
                        Row excelRow = sheet.createRow(rowNum++);
                        List<String> rowData = new ArrayList<>();

                        for (RectangularTextContainer cell : row) {
                            String cellText = cell.getText().trim();
                            cellText = cellText.replaceAll("[\r\n]+", " ")
                                               .replaceAll("\u00A0", " ")
                                               .replaceAll("\\s+", " ");

                            rowData.add(cellText);
                        }
                        while (rowData.size() < header.length) {
                            rowData.add("");
                        }
                        for (int i = 0; i < rowData.size(); i++) {
                            Cell cell = excelRow.createCell(i);
                            cell.setCellValue(rowData.get(i));
                        }
                    }
                }
            }


            for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(fileOut);
            System.out.println("✅ Arquivo Excel gerado com cabeçalho formatado: " + excelFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}