package com.example.excelutilizer.v1.service;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.http.server.reactive.ServerHttpResponse;
import reactor.core.publisher.Mono;

public interface ExcelService {

    Mono<Void> downloadExcel(ServerHttpResponse response, String filename, JsonNode jsonNode, Class<?> dtoClass);
}
