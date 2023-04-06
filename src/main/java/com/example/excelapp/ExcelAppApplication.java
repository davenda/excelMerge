package com.example.excelapp;

import org.joda.time.DateTime;
import org.joda.time.chrono.EthiopicChronology;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.telegram.telegrambots.meta.TelegramBotsApi;
import org.telegram.telegrambots.meta.exceptions.TelegramApiException;
import org.telegram.telegrambots.updatesreceivers.DefaultBotSession;

import javax.annotation.PostConstruct;
import java.util.Locale;

@SpringBootApplication
public class ExcelAppApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelAppApplication.class, args);
    }

    @PostConstruct
    public void registerBot(){
       try {
           System.out.println("Bot Starting");
           TelegramBotsApi botsApi = new TelegramBotsApi(DefaultBotSession.class);
           botsApi.registerBot(new BotService());
           System.out.println("Bot Started");
       } catch (TelegramApiException e) {
           e.printStackTrace();
       }
    }


}
