package com.example.transmission.controller;

import com.example.transmission.service.TransmissionChannelIncidentService;
import com.example.transmission.service.WordReportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/api/export")
public class TransmissionChannelController {
    @Autowired
    TransmissionChannelIncidentService transmissionChannelIncidentService;
    
    @Autowired
    WordReportService wordReportService;
    @GetMapping("/transmission-updated")
    public ResponseEntity<byte[]> downloadTransmissionUpdated(@RequestParam("date")
                                                                  @DateTimeFormat(pattern = "yyyy-MM-dd")
                                                                  LocalDate currentDate,
                                                              @RequestParam(name = "type", required = false) String type) {
        try {
            byte[] fileBytes = transmissionChannelIncidentService.exportExcelFile(currentDate);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=transmission_updated_v2.xlsx")
                    .contentType(MediaType.parseMediaType(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(fileBytes);
        } catch (Exception e) {
            // Log the error for debugging
            System.err.println("Error in downloadTransmissionUpdated() with cause = '" + 
                             (e.getCause() != null ? e.getCause().getClass().getSimpleName() : "NULL") + 
                             "' and exception = '" + e.getMessage() + "'");
            e.printStackTrace();
            
            // Return a more specific error response
            return ResponseEntity.internalServerError()
                    .body(("Error generating Excel file: " + e.getMessage()).getBytes());
        }
    }

    @GetMapping("/transmission-word-report")
    public ResponseEntity<byte[]> downloadWordReport(@RequestParam("date")
                                                       @DateTimeFormat(pattern = "yyyy-MM-dd")
                                                       LocalDate currentDate) {
        try {
            byte[] fileBytes = wordReportService.exportWordReport(currentDate);
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION,
                            "attachment; filename=bao_cao_truyen_so_lieu_" + 
                            currentDate.format(DateTimeFormatter.ofPattern("dd_MM_yyyy")) + ".docx")
                    .contentType(MediaType.parseMediaType(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                    .body(fileBytes);
        } catch (Exception e) {
            System.err.println("Error in downloadWordReport() with cause = '" + 
                             (e.getCause() != null ? e.getCause().getClass().getSimpleName() : "NULL") + 
                             "' and exception = '" + e.getMessage() + "'");
            e.printStackTrace();
            
            return ResponseEntity.internalServerError()
                    .body(("Error generating Word report: " + e.getMessage()).getBytes());
        }
    }
}
