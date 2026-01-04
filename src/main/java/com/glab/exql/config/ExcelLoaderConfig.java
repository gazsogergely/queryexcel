package com.glab.exql.config;

import com.glab.exql.service.ExcelToH2Service;
import org.springframework.boot.ApplicationRunner;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class ExcelLoaderConfig {

    @Bean
    ApplicationRunner loadExcel(ExcelToH2Service service) {
        return args -> service.loadExcelSheet("data.xlsx", "cica", "cica");
    }
}
