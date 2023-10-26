package com.example.demo.controller.export;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
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
/**
 * Controller pour réaliser l'export des clients.
 */
@Controller
@RequestMapping("export/clients")
public class ExportClientController {

    @Autowired
    private ClientService clientService;

    /**
     * Export des articles au format CSV.
     */
    @GetMapping("csv")
    public void exportCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();
        writer.println("Nom;Prénom;Age");
        List<ClientDto> clients = clientService.findAll();
        for (ClientDto client : clients) {
            writer.println(client.getNom() + ";" + client.getPrenom() + ";" + client.getAge());
        }
    }
    @GetMapping("xlsx")
    public void exportXLSX (HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.xlsx\"");
        OutputStream out = response.getOutputStream();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row row = sheet.createRow(0);
        Cell Hnom = row.createCell(0);
        Cell Hprenom = row.createCell(1);
        List<ClientDto> clients = clientService.findAll();
        for (int i=0;i<clients.size();i++) {
        	ClientDto client = clients.get(i);
        	Row row1 = sheet.createRow(i+1);
            Cell nom = row1.createCell(0);
            nom.setCellValue(client.getNom());
            Cell prenom = row1.createCell(1);
            prenom.setCellValue(client.getPrenom());
        }
        workbook.write(out);
    }

}
