package com.example.excelutilizer.v1.controller;

import com.example.excelutilizer.v1.dto.ExampleDto;
import com.example.excelutilizer.v1.service.ExcelService;
import com.fasterxml.jackson.databind.JsonNode;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.server.reactive.ServerHttpResponse;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import reactor.core.publisher.Mono;

@Slf4j
@RestController("/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;

    @PostMapping("/download")
    public Mono<Void> downloadExcel(@RequestParam String filename,
        @RequestBody JsonNode jsonNode,
        ServerHttpResponse response) {

        return excelService.downloadExcel(response, filename, jsonNode, ExampleDto.class);
    }
}
