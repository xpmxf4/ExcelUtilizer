package com.example.excelutilizer.v1.service;

import com.fasterxml.jackson.databind.JsonNode;
import java.io.ByteArrayOutputStream;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.core.io.buffer.DataBuffer;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.server.reactive.ServerHttpResponse;
import org.springframework.stereotype.Service;
import reactor.core.publisher.Mono;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {

    private final JsonToExcelService jsonToExcelService;

    @Override
    public Mono<Void> downloadExcel(ServerHttpResponse response, String filename, JsonNode jsonNode, Class<?> dtoClass) {
        return Mono.fromCallable(() -> {
            Workbook workbook = jsonToExcelService.convertJsonToExcel(jsonNode, dtoClass.getClass());
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();
            return outputStream.toByteArray();
        }).flatMap(bytes -> {
            response.getHeaders().set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\".xlsx");
            response.getHeaders().setContentType(MediaType.APPLICATION_OCTET_STREAM);
            DataBuffer buffer = response.bufferFactory().wrap(bytes);
            return response.writeWith(Mono.just(buffer));
        });
    }
}
