package com.example.excelapp;

import org.springframework.web.bind.annotation.*;

@RestController
public class BotController {


    private final BotService bot;

    public BotController(BotService bot) {
        this.bot = bot;
    }

    @RequestMapping(value = "/sendMessage", method = RequestMethod.GET)
    public void sendMessage(@RequestParam("chatId") String chatId, @RequestParam("text") String text) {
        // forward the message to your bot for processing
        bot.sendTextMessage(chatId, text);
    }
}
