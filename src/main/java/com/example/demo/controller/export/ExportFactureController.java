package com.example.demo.controller.export;

import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import com.example.demo.dto.ClientDto;
import com.example.demo.service.ClientService;

import java.io.OutputStream;
import java.util.Arrays;
import org.springframework.web.bind.annotation.PathVariable;
import com.example.demo.dto.FactureDto;
import com.example.demo.dto.LigneFactureDto;
import com.example.demo.service.FactureService;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;


/**
 * Controller pour réaliser l'export des factures.
 */
@Controller
@RequestMapping("export/factures")
public class ExportFactureController {
	@Autowired
    private FactureService factureservice;
    @GetMapping("{idFacture}/pdf")
    public void exportPDF(@PathVariable Long idFacture, HttpServletRequest request, HttpServletResponse response)
        throws IOException, DocumentException {
        response.setHeader("Content-Disposition", "attachment; filename=\"export-facture-" + idFacture + ".pdf\"");
        OutputStream outputStream = response.getOutputStream();
        Document document = new Document();
        PdfWriter writer = PdfWriter.getInstance(document, outputStream);
        document.open();
        List<FactureDto> factures = factureservice.findAll();
        for(FactureDto facture: factures) {
        	if(facture.getId().equals(idFacture)) {
                Paragraph paragrapheHeader1 = new Paragraph();
                paragrapheHeader1.add("hello");
                document.add(paragrapheHeader1);

                PdfPTable table = new PdfPTable(4);
                table.addCell("Article");
                table.addCell("Prix unitaire");
                table.addCell("Quantité");
                table.addCell("Prix");
                List<LigneFactureDto> ClientFactures = facture.getLigneFactures();
                for(LigneFactureDto ClientFacture:ClientFactures) {
                	table.addCell(ClientFacture.getDesignation().toString());
                    table.addCell(ClientFacture.getPrixUnitaire().toString());
                    table.addCell(ClientFacture.getQuantite().toString());
                    table.addCell(ClientFacture.getPrix().toString());
                }
                document.add(table);
                document.close();
        	}
        }
    }
    @GetMapping("xlsx")
    public void exportXLSX(HttpServletRequest request, HttpServletResponse response)
        throws IOException, DocumentException {
    	response.setHeader("Content-Disposition", "attachment; filename=\"export-facture.xlsx\"");
        ServletOutputStream outputStream = response.getOutputStream();
        Workbook workbook = new XSSFWorkbook();
        List<FactureDto> factures = factureservice.findAll();
        for(FactureDto facture: factures) {
        	ClientDto client = facture.getClient();
        	String nomprenom = client.getNom()+" "+ client.getPrenom();
        	if(workbook.getSheetIndex(nomprenom)<0) {
        		Sheet sheetClientF = workbook.createSheet(nomprenom);
        		addcell(sheetClientF,0, Arrays.asList(
                		"Prenom : ",
                		client.getPrenom()));
                addcell(sheetClientF,1,Arrays.asList(
                		"année de naissence",
                		Integer.toString(client.getAge())));
                List<LigneFactureDto> ClientFactures = facture.getLigneFactures();
                Row rowfacture = sheetClientF.createRow(2);
                rowfacture.createCell(0).setCellValue(ClientFactures.size()+" Facture(s)");
            	Sheet sheetfacture = workbook.createSheet("Factures N°"+facture.getId());
            	addcell(sheetfacture, 0, Arrays.asList(
            			"Désignation",
            			"Quantité",
            			"Prix unitaire",
            			"prix total"
            			));
                for(int i=0;i<ClientFactures.size();i++) {
                	LigneFactureDto ClientFacture = ClientFactures.get(i);
                	int irow = i+1;
                	addcell(sheetfacture, irow, Arrays.asList(
                			ClientFacture.getDesignation(),
                			ClientFacture.getQuantite().toString(),
                			ClientFacture.getPrix().toString().replace('.', ',')
                			));
            		sheetfacture.getRow(irow).createCell(3).setCellFormula("B"+(irow+1)+"*C"+(irow+1));
                }
        	}
    	}
        workbook.write(outputStream);
        workbook.close();

    }
	private void addcell(Sheet sheet, int row, List<String> cells) {
		Row rowprenomClient = sheet.createRow(row);
		for(int i=0;i<cells.size();i++) {
			String cell = cells.get(i);
	        rowprenomClient.createCell(i).setCellValue(cell);
		}
	}
}